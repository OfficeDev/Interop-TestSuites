namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Notification ROPs.
    /// </summary>
    [TestClass]
    public class S12_NotificationROPs : TestSuiteBase
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
        /// This method tests the ROP buffers of RopRegisterNotification, RopPending and RopNotify.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S12_TC01_TestRopPending()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopRegisterNotification request and verify the success response.
            #region RopRegisterNotification success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopRegisterNotificationRequest registerNotificationRequest;
            RopRegisterNotificationResponse registerNotificationResponse;

            registerNotificationRequest.RopId = (byte)RopId.RopRegisterNotification;

            registerNotificationRequest.LogonId = TestSuiteBase.LogonId;
            registerNotificationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            registerNotificationRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex0;

            // The server MUST send notifications to the client when CriticalError events occur within the scope of interest
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.NewMail;
            registerNotificationRequest.Reserved = TestSuiteBase.Reserved;

            // TRUE: the scope for notifications is the entire database
            registerNotificationRequest.WantWholeStore = TestSuiteBase.NonZero;

            registerNotificationRequest.FolderId = logonResponse.FolderIds[4];
            registerNotificationRequest.MessageId = MS_OXCROPSAdapter.MessageIdForRops;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x02.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x04, which means the server sends notifications to the client when ObjectCreated events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.ObjectCreated;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x04.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x08, which means the server sends notifications to the client when ObjectDeleted events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.ObjectDeleted;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x08.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x10, which means the server sends notifications to the client when ObjectModified events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.ObjectModified;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x10.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x20, which means the server sends notifications to the client when ObjectMoved events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.ObjectMoved;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x20.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x40, which means the server sends notifications to the client when ObjectCopied events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.ObjectCopied;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x40.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Set NotificationTypes to 0x80, which means the server sends notifications to the client when SearchCompleted events occur
            // within the scope of interest, as specified in [MS-OXCNOTIF].
            registerNotificationRequest.NotificationTypes = (byte)NotificationTypes.SearchCompleted;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopRegisterNotification request:NotificationTypes=0x80.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                registerNotificationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            registerNotificationResponse = (RopRegisterNotificationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 2: Create message,Save message to test RopPending, RopNotify and RopBufferTooSmall
            #region RopRegisterNotification success response

            #region prepare rops for createmessage, savemessage and release
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest();
            RopReleaseRequest releaseRequest = new RopReleaseRequest();
            this.PrepareRops(logonResponse, ref createMessageRequest, ref saveChangesMessageRequest, ref releaseRequest);
            #endregion
            // Totally create message loop count.
            int loopCount;
            uint tableHandle = 0;

            string transportSeq = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower();
            if (transportSeq == "mapi_http")
            {
                loopCount = 20;
                this.CreateVastMessages(ref logonResponse, out tableHandle, loopCount, createMessageRequest, saveChangesMessageRequest, releaseRequest);
            }
            else
            {
                loopCount = 1000;
                this.CreateSingleProcessEachLoop(ref logonResponse, out tableHandle, loopCount, createMessageRequest, saveChangesMessageRequest, releaseRequest);
            }

            #region RopBufferTooSmall response

            List<ISerializable> ropRequests1 = new List<ISerializable>();
            List<uint> inputObjects1 = new List<uint>
            {
                this.inputObjHandle
            };
            List<IDeserializable> ropResponses1 = new List<IDeserializable>();
            RopGetPropertiesAllRequest getPropertiesAllRequest;

            getPropertiesAllRequest.RopId = (byte)RopId.RopGetPropertiesAll;
            getPropertiesAllRequest.LogonId = TestSuiteBase.LogonId;
            getPropertiesAllRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned.
            getPropertiesAllRequest.PropertySizeLimit = TestSuiteBase.PropertySizeLimit;

            getPropertiesAllRequest.WantUnicode = Convert.ToUInt16(TestSuiteBase.Zero);
            byte count = 255;
            for (byte i = 1; i < count; i++)
            {
                ropRequests1.Add(getPropertiesAllRequest);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the ProcessMutipleRops request to verify RopBufferTooSmall response.");

            // Verify the RopBufferTooSmall
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(ropRequests1, inputObjects1, ref ropResponses1, ref this.rawData, RopResponseType.FailureResponse);

            #endregion

            List<IDeserializable> ropResponses4 = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to resubmit all the GetPropertiesAll request.");

            // All the requests are GetPropertiesAllRequest, resubmit the request in another call and get the responses in the ropResponses4
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(ropRequests1, inputObjects1, ref ropResponses4, ref this.rawData, RopResponseType.SuccessResponse);

            #region RopSetColumns success response

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            PropertyTag[] propertyTags = CreateSampleContentsTablePropertyTags();
            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = TestSuiteBase.LogonId;
            setColumnsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSetColumns request.");

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

            #region RopQueryRows success response: send RopSetColumns and RopQueryRows in a request buffer
            RopQueryRowsRequest queryRowsRequest;

            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = TestSuiteBase.LogonId;
            queryRowsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;
            queryRowsRequest.ForwardRead = TestSuiteBase.NonZero;

            // Set RowCount to 0x01, which the number of requested rows, as specified in [MS-OXCROPS].
            queryRowsRequest.RowCount = TestSuiteBase.RowCount;

            List<ISerializable> ropRequests = new List<ISerializable>
            {
                setColumnsRequest, queryRowsRequest
            };
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex2;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex3;
            createMessageRequest.FolderId = logonResponse.FolderIds[4];
            ropRequests.Add(createMessageRequest);
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex3;
            ropRequests.Add(saveChangesMessageRequest);
            List<uint> inputObjects = new List<uint>
            {
                tableHandle,
                0xFFFF,
                this.inputObjHandle,
                0xFFFF
            };

            // 0xFFFF indicates a default input handle.
            List<IDeserializable> ropResponses = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the multiple ROPs request to verify the RopNotify response.");

            // Verify the RopNotify
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(ropRequests, inputObjects, ref ropResponses, ref this.rawData, RopResponseType.SuccessResponse);

            // Set RowCount to 0x0600, which the number of requested rows, as specified in [MS-OXCROPS].
            queryRowsRequest.RowCount = TestSuiteBase.RowCount;

            List<ISerializable> ropRequests2 = new List<ISerializable>
            {
                setColumnsRequest, queryRowsRequest
            };
            createMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex2;
            createMessageRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex3;
            createMessageRequest.FolderId = logonResponse.FolderIds[4];
            ropRequests2.Add(createMessageRequest);
            saveChangesMessageRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex3;
            ropRequests2.Add(saveChangesMessageRequest);
            List<uint> inputObjects2 = new List<uint>
            {
                tableHandle,
                0xFFFF,
                this.inputObjHandle,
                0xFFFF
            };

            // 0xFFFF indicates a default input handle.
            List<IDeserializable> ropResponses2 = new List<IDeserializable>();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the multiple ROPs request to verify the RopPending response.");

            // Verify the RopPending
            this.responseSOHs = cropsAdapter.ProcessMutipleRops(ropRequests2, inputObjects2, ref ropResponses2, ref this.rawData, RopResponseType.SuccessResponse);

            bool isContainedRopPending = false;
            for (int i = 1; i < ropResponses2.Count; i++)
            {
                if (ropResponses2[i] is RopPendingResponse)
                {
                    isContainedRopPending = true;
                    break;
                }
            }

            // Send an empty ROP to verify all queued RopNotify response have got in last step.
            List<ISerializable> ropRequestsEmpty = new List<ISerializable>(0x02);
            List<IDeserializable> ropResponsesNull = new List<IDeserializable>();

            this.responseSOHs = cropsAdapter.ProcessMutipleRops(ropRequestsEmpty, inputObjects, ref ropResponsesNull, ref this.rawData, RopResponseType.SuccessResponse);

            if (ropResponsesNull.Count == 0)
            {
                if (Common.IsRequirementEnabled(469303, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R469303");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R469303
                    Site.CaptureRequirementIfIsTrue(
                        isContainedRopPending,
                        469303,
                        @"[In Appendix B: Product Behavior] Implementation does include a RopPending ROP response (section 2.2.14.3) even though the ROP output buffer contains all queued RopNotify ROP responses (section 2.2.14.2). (<16> Section 3.1.5.1.3: Exchange 2007 follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(4693031, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4693031");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R4693031
                    Site.CaptureRequirementIfIsFalse(
                        isContainedRopPending,
                        4693031,
                        @"[In Appendix B: Product Behavior] Implementation does not include a RopPending ROP response if the ROP output buffer contain all queued RopNotify ROP responses. (Exchange 2010 and above follow this behavior.)");
                }
            }
            #endregion

            #endregion
        }

        /// <summary>
        /// This method tests the RPC fail error code while exceeding the max PcbOut.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S12_TC02_TestFailRPCForMaxPcbOut()
        {
            this.CheckTransportIsSupported();

            if (Common.IsRequirementEnabled(454509, this.Site)
                || Common.IsRequirementEnabled(20009, this.Site))
            {
                this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

                // Define data for properties
                string dataForProperties = string.Empty;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;
                dataForProperties += TestSuiteBase.DataForProperties;

                int loopCounter = (LoopCounter / Encoding.ASCII.GetBytes(dataForProperties).Length) + 1;

                #region Common operations for RopGetStreamSize,RopSetStreamSize and RopSeekStream

                // Log on to a private mailbox.
                RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Call GetCreatedMessageHandle method to create a message and get its handle.");

                // Call GetCreatedMessageHandle method to create a message and get its handle.
                uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

                // Call GetOpenedStreamHandle method to open stream and get its handle.
                uint streamObjectHandle;

                for (int i = 0; i < loopCounter; i++)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Call GetOpenedStreamHandle method to open stream and get its handle:loop counter i={0}", i);

                    streamObjectHandle = this.GetOpenedStreamHandle(
                        messageHandle,
                        (ushort)(this.propertyDictionary[PropertyNames.UserSpecified].PropertyId + i));

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Call WriteStream method to write stream:loop counter i={0}", i);

                    // Call WriteStream method to write stream.
                    this.WriteStream(streamObjectHandle, dataForProperties);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Call CommitStream method to commit source stream:loop counter i={0}", i);

                    // Call CommitStream method to commit source stream.
                    this.CommitStream(streamObjectHandle);

                    #region Release the stream.
                    RopReleaseRequest releaseRequest;

                    releaseRequest.RopId = (byte)RopId.RopRelease;
                    releaseRequest.LogonId = TestSuiteBase.LogonId;

                    // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                    // where the handle for the input Server object is stored.
                    releaseRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopRelease request.");

                    // Send a RopRelease request to release all resources associated with the Server object.
                    this.responseSOHs = cropsAdapter.ProcessSingleRop(
                        releaseRequest,
                        streamObjectHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    #endregion
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Call SaveMessage method to save message.");

                // Call SaveMessage method to save message.
                this.SaveMessage(messageHandle);

                #endregion

                #region Call RopGetPropertiesAll to make the response buffer larger than the maximum size specified in pcbOut

                RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest
                {
                    RopId = (byte)RopId.RopGetPropertiesAll,
                    LogonId = TestSuiteBase.LogonId,
                    InputHandleIndex = TestSuiteBase.InputHandleIndex0,
                    PropertySizeLimit = TestSuiteBase.PropertySizeLimit,
                    WantUnicode = (ushort)Zero
                };

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopGetPropertiesAll request.");

                // Send the RopGetPropertiesAll request and verify RPC error response.
                this.responseSOHs = cropsAdapter.ProcessSingleRopWithOptionResponseBufferSize(
                    getPropertiesAllRequest,
                    messageHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.RPCError,
                    0x10);

                #endregion
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only under Exchange 2010 and above, since Exchange 2007 does not support fail the RPC with 0x0000047D if the ROP output buffer is larger than the maximum size specified in request.");
            }
        }

        #endregion

        #region Common methods and property

        /// <summary>
        /// Get Opened Stream Handle.
        /// </summary>
        /// <param name="messageHandle">the message handle</param>
        /// <param name="propertyId">The ID of property</param>
        /// <returns>Return the opened stream handle</returns>
        private uint GetOpenedStreamHandle(uint messageHandle, ushort propertyId)
        {
            RopOpenStreamRequest openStreamRequest;
            RopOpenStreamResponse openStreamResponse;

            PropertyTag tag;
            tag.PropertyId = propertyId;
            tag.PropertyType = (ushort)PropertyType.PtypString;

            openStreamRequest.RopId = (byte)RopId.RopOpenStream;
            openStreamRequest.LogonId = TestSuiteBase.LogonId;
            openStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openStreamRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openStreamRequest.PropertyTag = tag;
            openStreamRequest.OpenModeFlags = (byte)StreamOpenModeFlags.Create;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopOpenStream request in GetOpenedStreamHandle method.");

            // Send the RopOpenStream request and verify success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openStreamRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openStreamResponse = (RopOpenStreamResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
            uint streamObjectHandle = responseSOHs[0][openStreamResponse.OutputHandleIndex];
            return streamObjectHandle;
        }

        /// <summary>
        /// Write Stream.
        /// </summary>
        /// <param name="streamHandle">the opened stream handle</param>
        /// <param name="dataString">The string to be written</param>
        private void WriteStream(uint streamHandle, string dataString)
        {
            RopWriteStreamRequest writeStreamRequest;
            RopWriteStreamResponse writeStreamResponse;

            byte[] data = Encoding.ASCII.GetBytes(dataString + "\0");
            writeStreamRequest.RopId = (byte)RopId.RopWriteStream;
            writeStreamRequest.LogonId = TestSuiteBase.LogonId;
            writeStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            writeStreamRequest.DataSize = (ushort)data.Length;
            writeStreamRequest.Data = data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopWriteStream request in WriteStream method.");

            // Send the RopWriteStream request and verify success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                writeStreamRequest,
                streamHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            writeStreamResponse = (RopWriteStreamResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                writeStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");
        }

        #endregion
    }
}