namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Stream ROPs. 
    /// </summary>
    [TestClass]
    public class S07_StreamROPs : TestSuiteBase
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
        /// This method tests the ROP buffers of RopOpenStream, RopReadStream, 
        /// RopWriteStream, RopCommitStream and RopWriteAndCommitStream.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S07_TC01_TestRopsOpenReadWriteCommitStream()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create message and then open a stream.
            #region Common operations for RopReadStream,RopWriteStream,RopCommitStream and RopWriteAndCommitStream

            // Common variable for RopWriteStream and RopWriteAndCommitStream.
            byte[] data = Encoding.ASCII.GetBytes(SampleStreamData + "\0");

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedMessageHandle method to create a message.");

            // Create a message.
            uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Test RopOpenStream success response.
            #region Test RopOpenStream success response

            RopOpenStreamRequest openStreamRequest;
            RopOpenStreamResponse openStreamResponse;

            // Client defines a new property.
            PropertyTag tag;
            tag.PropertyId = TestSuiteBase.UserDefinedPropertyId;
            tag.PropertyType = (ushort)PropertyType.PtypString;

            openStreamRequest.RopId = (byte)RopId.RopOpenStream;

            openStreamRequest.LogonId = TestSuiteBase.LogonId;
            openStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openStreamRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openStreamRequest.PropertyTag = tag;
            openStreamRequest.OpenModeFlags = (byte)StreamOpenModeFlags.Create;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenStream request.");

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

            #endregion

            // Test RopOpenStream failure response.
            #region Test RopOpenStream failure response

            openStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenStream request.");

            // Send the RopOpenStream request and verify failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openStreamRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            openStreamResponse = (RopOpenStreamResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openStreamResponse.ReturnValue,
                "if ROP failure, the ReturnValue of its response is not 0(success)");

            #endregion

            #endregion

            // Step 2: Verify RopReadStream Response when ByteCount is not equal to 0xBABE.
            #region Verify RopReadStream Response when ByteCount is not equal to 0xBABE.

            RopReadStreamRequest readStreamRequest;
            RopReadStreamResponse readStreamResponse;

            readStreamRequest.RopId = (byte)RopId.RopReadStream;
            readStreamRequest.LogonId = TestSuiteBase.LogonId;
            readStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ByteCount to a value other than 0xBABE, then verify the response.
            readStreamRequest.ByteCount = TestSuiteBase.ByteCountForRopReadStream;

            readStreamRequest.MaximumByteCount = TestSuiteBase.MaximumByteCount;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopReadStream request.");

            // Send the RopReadStream request and verify RopReadStream Response when ByteCount is not equal to 0xBABE.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            readStreamResponse = (RopReadStreamResponse)response;

            if (readStreamRequest.ByteCount != TestSuiteBase.ByteCount)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3227,the ByteCount:{0},the DataSize:{1}", readStreamRequest.ByteCount, readStreamResponse.DataSize);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3227
                // The maximum size is specified in the request buffer ByteCount field, must greater or equal than response DataSize field.
                bool isVerifyR3227 = readStreamRequest.ByteCount >= readStreamResponse.DataSize;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR3227,
                    3227,
                    @"[In RopReadStream ROP Response Buffer,DataSize (2 bytes),The maximum size is specified in the request buffer by one of the following:]The ByteCount field, when the value of the ByteCount value is not equal to 0xBABE.");
            }

            #endregion

            // Step 3: Verify RopReadStream Response when ByteCount is equal to 0xBABE.
            #region Verify RopReadStream Response when ByteCount is equal to 0xBABE.

            // Set ByteCount to 0xBABE, then verify the response.
            readStreamRequest.ByteCount = TestSuiteBase.ByteCount;
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            readStreamResponse = (RopReadStreamResponse)response;

            if (readStreamRequest.ByteCount == TestSuiteBase.ByteCount)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3228,the MaximumByteCount:{0},the DataSize:{1}", readStreamRequest.MaximumByteCount, readStreamResponse.DataSize);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3228
                // The maximum size is specified in the request buffer MaximumByteCount field, must greater or equal than response DataSize field.
                bool isVerifyR3228 = readStreamRequest.MaximumByteCount >= readStreamResponse.DataSize;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR3228,
                    3228,
                    @"[In RopReadStream ROP Response Buffer,DataSize (2 bytes),The maximum size is specified in the request buffer by one of the following:]The MaximumByteCount field, when the value of the ByteCount field is equal to 0xBABE.");
            }
            #endregion

            // Step 4: Verify RopReadStream Response when MaximumByteCount is larger than 0x80000000.
            #region Verify RopReadStream Response when MaximumByteCount is larger than 0x80000000

            // Set MaximumByteCount to be larger than 0x80000000.
            readStreamRequest.MaximumByteCount = TestSuiteBase.ExceedMaxCount;

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.RPCError);

            #endregion

            // Step 5: Send the RopWriteStream request and verify the success response.
            #region RopWriteStream Response

            RopWriteStreamRequest writeStreamRequest;
            RopWriteStreamResponse writeStreamResponse;

            writeStreamRequest.RopId = (byte)RopId.RopWriteStream;
            writeStreamRequest.LogonId = TestSuiteBase.LogonId;
            writeStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            writeStreamRequest.DataSize = (ushort)data.Length;
            writeStreamRequest.Data = data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopWriteStream request.");

            // Send the RopWriteStream request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                writeStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            writeStreamResponse = (RopWriteStreamResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                writeStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 6: Send the RopCommitStream request and verify the success response.
            #region RopCommitStream Response

            RopCommitStreamRequest commitStreamRequest;
            RopCommitStreamResponse commitStreamResponse;

            commitStreamRequest.RopId = (byte)RopId.RopCommitStream;
            commitStreamRequest.LogonId = TestSuiteBase.LogonId;
            commitStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopCommitStream request.");

            // Send the RopCommitStream request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                commitStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            commitStreamResponse = (RopCommitStreamResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                commitStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 7: Send the RopWriteAndCommitStream request and verify the success response.
            #region RopWriteAndCommitStream response

            RopWriteAndCommitStreamRequest writeAndCommitStreamRequest;

            writeAndCommitStreamRequest.RopId = (byte)RopId.RopWriteAndCommitStream;
            writeAndCommitStreamRequest.LogonId = TestSuiteBase.LogonId;
            writeAndCommitStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            writeAndCommitStreamRequest.DataSize = (ushort)data.Length;
            writeAndCommitStreamRequest.Data = data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopWriteAndCommitStream request.");

            // Send the RopWriteAndCommitStream request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                writeAndCommitStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            // Refer to MS-OXCPRPT: Exchange 2003 and Exchange 2007 implement the RopWriteAndCommitStream ROP.
            if (Common.IsRequirementEnabled(752001, this.Site))
            {
                // Because the response of RopWriteAndCommitStream is the same to RopWriteStream.
                RopWriteStreamResponse writeAndCommitStreamResponse;
                writeAndCommitStreamResponse = (RopWriteStreamResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R752001");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R752001
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    writeAndCommitStreamResponse.ReturnValue,
                    752001,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopWriteAndCommitStream ROP. (Exchange 2007 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    writeAndCommitStreamResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");
            }

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopGetStreamSize, RopSetStreamSize and RopSeekStream.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S07_TC02_TestRopsGetSetStreamSizeAndSeekStream()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create message and get its handle, open a stream and get its handle.
            #region Common operations for RopGetStreamSize,RopSetStreamSize and RopSeekStream

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedMessageHandle method to create a message and get its handle.");

            // Call GetCreatedMessageHandle method to create a message and get its handle.
            uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetOpenedStreamHandle method to open stream and get its handle.");

            // Call GetOpenedStreamHandle method to open stream and get its handle.
            uint streamObjectHandle = this.GetOpenedStreamHandle(messageHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call WriteStream method to write stream.");

            // Call WriteStream method to write stream.
            this.WriteStream(streamObjectHandle);

            #endregion

            // Step 2: Send the RopGetStreamSize request and verify the success response.
            #region RopGetStreamSize success response

            RopGetStreamSizeRequest getStreamSizeRequest;
            RopGetStreamSizeResponse getStreamSizeResponse;

            getStreamSizeRequest.RopId = (byte)RopId.RopGetStreamSize;

            getStreamSizeRequest.LogonId = TestSuiteBase.LogonId;
            getStreamSizeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetStreamSize request.");

            // Send the RopGetStreamSize request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getStreamSizeRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getStreamSizeResponse = (RopGetStreamSizeResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getStreamSizeResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 3: Send the RopGetStreamSize request and verify the failure response.
            #region RopGetStreamSize failure response

            getStreamSizeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetStreamSize request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getStreamSizeRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getStreamSizeResponse = (RopGetStreamSizeResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getStreamSizeResponse.ReturnValue,
                "if ROP failure, the ReturnValue of its response is not 0(success)");

            #endregion

            // Step 4: Send the RopSetStreamSize request and verify the success response.
            #region RopSetStreamSize Response

            RopSetStreamSizeRequest setStreamSizeRequest;
            RopSetStreamSizeResponse setStreamSizeResponse;

            setStreamSizeRequest.RopId = (byte)RopId.RopSetStreamSize;
            setStreamSizeRequest.LogonId = TestSuiteBase.LogonId;
            setStreamSizeRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Change the original stream size to 0x00000000000000FF.
            setStreamSizeRequest.StreamSize = TestSuiteBase.StreamSize;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetStreamSize request.");

            // Send the RopSetStreamSize request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setStreamSizeRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setStreamSizeResponse = (RopSetStreamSizeResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setStreamSizeResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 5: Send the RopSeekStream request and verify the success response.
            #region RopSeekStream success response

            RopSeekStreamRequest seekStreamRequest;
            RopSeekStreamResponse seekStreamResponse;

            seekStreamRequest.RopId = (byte)RopId.RopSeekStream;
            seekStreamRequest.LogonId = TestSuiteBase.LogonId;
            seekStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            seekStreamRequest.Origin = (byte)Origin.Beginning;

            // Defined by tester, less than the stream size.
            seekStreamRequest.Offset = TestSuiteBase.Offset;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSeekStream request.");

            // Send the RopSeekStream request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            seekStreamResponse = (RopSeekStreamResponse)response;
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                seekStreamResponse.ReturnValue,
                "if ROP succeeds, the ReturnValue of its response is 0(success)");

            #endregion

            // Step 6: Send the RopSeekStream request and verify the failure response.
            #region RopSeekStream failure response

            seekStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopSeekStream request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                seekStreamRequest,
                streamObjectHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            seekStreamResponse = (RopSeekStreamResponse)response;
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                seekStreamResponse.ReturnValue,
                "if ROP failure, the ReturnValue of its response is not 0(success)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopLockRegionStream and RopUnlockRegionStream.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S07_TC03_TestRopsLockAndUnlockRegionStream()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create message and get its handle, open stream and get its handle.
            #region Common operations for RopLockRegionStream and RopUnlockRegionStream

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedMessageHandle method to create message and get its handle.");

            // Call GetCreatedMessageHandle method to create message and get its handle.
            uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetOpenedStreamHandle method to open stream and get its handle.");

            // Call GetOpenedStreamHandle method to open stream and get its handle.
            uint streamObjectHandle = this.GetOpenedStreamHandle(messageHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call WriteStream method to write stream.");

            // Call WriteStream method to write stream.
            this.WriteStream(streamObjectHandle);

            #endregion

            // Refer to MS-OXCPRPT: Exchange 2003 and Exchange 2007 implement the RopLockRegionStream ROP.
            if (Common.IsRequirementEnabled(750001, this.Site))
            {
                // Step 2: Send the RopLockRegionStream request and verify the success response.
                #region RopLockRegionStream Response

                RopLockRegionStreamRequest lockRegionStreamRequest = new RopLockRegionStreamRequest();
                RopLockRegionStreamResponse lockRegionStreamResponse;

                lockRegionStreamRequest.RopId = (byte)RopId.RopLockRegionStream;

                lockRegionStreamRequest.LogonId = TestSuiteBase.LogonId;
                lockRegionStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
                lockRegionStreamRequest.RegionOffset = TestSuiteBase.RegionOffset;

                // Defined by tester, less than the stream size.
                lockRegionStreamRequest.RegionSize = TestSuiteBase.RegionSize;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopLockRegionStream request.");

                // Send the RopLockRegionStream request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    lockRegionStreamRequest,
                    streamObjectHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                lockRegionStreamResponse = (RopLockRegionStreamResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R750001");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R750001
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    lockRegionStreamResponse.ReturnValue,
                    750001,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopLockRegionStream ROP. (Exchange 2007 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    lockRegionStreamResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion
            }

            // Refer to MS-OXCPRPT: Exchange 2003 and Exchange 2007 implement the RopUnlockRegionStream ROP.
            if (Common.IsRequirementEnabled(751001, this.Site))
            {
                // Step 3: Send the RopUnlockRegionStream request and verify the success response.
                #region RopUnlockRegionStream response

                RopUnlockRegionStreamRequest unlockRegionStreamRequest = new RopUnlockRegionStreamRequest();
                RopUnlockRegionStreamResponse unlockRegionStreamResponse;

                unlockRegionStreamRequest.RopId = (byte)RopId.RopUnlockRegionStream;
                unlockRegionStreamRequest.LogonId = TestSuiteBase.LogonId;
                unlockRegionStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Beginning of the stream.
                unlockRegionStreamRequest.RegionOffset = TestSuiteBase.RegionOffset;

                // Defined by tester, less than the stream size.
                unlockRegionStreamRequest.RegionSize = TestSuiteBase.RegionSize;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopUnlockRegionStream request.");

                // Step 4: Send the RopUnlockRegionStream request and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    unlockRegionStreamRequest,
                    streamObjectHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                unlockRegionStreamResponse = (RopUnlockRegionStreamResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R751001");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R751001
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    unlockRegionStreamResponse.ReturnValue,
                    751001,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopUnlockRegionStream ROP. (Exchange 2007 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    unlockRegionStreamResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion
            }
        }

        /// <summary>
        /// This method tests the ROP buffers of RopCloneStream and RopCopyToStream.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S07_TC04_TestRopsCloneAndCopyToStream()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 5: Preparations for RopCloneStream and RopCopyToStream.
            #region Common operations for RopCloneStream and RopCopyToStream

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call GetCreatedMessageHandle method to create message and get its handle.");

            // Call GetCreatedMessageHandle method to create message and get its handle.
            uint sourceMessageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call GetOpenedStreamHandle method to open stream and get its handle.");

            // Call GetOpenedStreamHandle method to open stream and get its handle.
            uint sourceStreamObjectHandle = this.GetOpenedStreamHandle(sourceMessageHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call WriteStream method to write stream.");

            // Call WriteStream method to write stream.
            this.WriteStream(sourceStreamObjectHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call CommitStream method to commit source stream.");

            // Call CommitStream method to commit source stream.
            this.CommitStream(sourceStreamObjectHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call SaveMessage method to save message.");

            // Call SaveMessage method to save message.
            this.SaveMessage(sourceMessageHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call GetOpenedStreamHandle method to open the source stream again.");

            // Open the source stream again.
            sourceStreamObjectHandle = this.GetOpenedStreamHandle(sourceMessageHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call GetCreatedMessageHandle method to create a second message.");

            // Create a second message.
            uint destinationMessageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5:Call GetOpenedStreamHandle method to open stream.");

            // Open stream, used as destination stream.
            uint destinationStreamObjectHandle = this.GetOpenedStreamHandle(destinationMessageHandle);

            // Copy stream, from source to destination.
            List<uint> handleList = new List<uint>
            {
                sourceStreamObjectHandle, destinationStreamObjectHandle
            };

            #endregion

            // Refer to MS-OXCPRPT: Exchange 2003 and Exchange 2007 implement the RopCloneStream ROP.
            if (Common.IsRequirementEnabled(753001, this.Site))
            {
                // Step 6: Send the RopCloneStream request and verify the success response.
                #region RopCloneStream response

                RopCloneStreamRequest cloneStreamRequest;
                RopCloneStreamResponse cloneStreamResponse;

                cloneStreamRequest.RopId = (byte)RopId.RopCloneStream;
                cloneStreamRequest.LogonId = TestSuiteBase.LogonId;
                cloneStreamRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
                cloneStreamRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopCloneStream request.");

                // Send the RopCloneStream request and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                    cloneStreamRequest,
                    handleList,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                cloneStreamResponse = (RopCloneStreamResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R753001");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R753001
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    cloneStreamResponse.ReturnValue,
                    753001,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopCloneStream ROP. (Exchange 2007 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    cloneStreamResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion
            }

            // Refer to MS-OXCPRPT: The initial release version of Exchange 2010 does not implement the RopCopyToStream ROP.
            if (Common.IsRequirementEnabled(8670901, this.Site))
            {
                // Step 7: Send the RopCopyToStream request and verify the success response.
                #region RopCopyToStream success response

                RopCopyToStreamRequest copyToStreamRequest;
                RopCopyToStreamResponse copyToStreamResponse;

                copyToStreamRequest.RopId = (byte)RopId.RopCopyToStream;
                copyToStreamRequest.LogonId = TestSuiteBase.LogonId;
                copyToStreamRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;
                copyToStreamRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;

                // Set ByteCount to a value less than the length of original property.
                copyToStreamRequest.ByteCount = TestSuiteBase.ByteCountForRopCopyToStream;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopCopyToStream request.");

                // Send the RopCopyToStream request and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                    copyToStreamRequest,
                    handleList,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                copyToStreamResponse = (RopCopyToStreamResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R8670901");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R8670901
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    copyToStreamResponse.ReturnValue,
                    8670901,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopCopyToStream ROP. (Exchange 2007 and Exchange 2013 follows this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    copyToStreamResponse.ReturnValue,
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                #endregion

                // Step 8: Send the RopCopyToStream request and verify the null destination failure response.
                #region RopCopyToStream null destination failure response

                handleList.RemoveAt(1);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopCopyToStream request.");

                this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                    copyToStreamRequest,
                    handleList,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.NullDestinationFailureResponse);
                copyToStreamResponse = (RopCopyToStreamResponse)response;
                Site.Assert.AreEqual<uint>(
                    MS_OXCROPSAdapter.ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                    copyToStreamResponse.ReturnValue,
                    "if ROP null destination failure, the ReturnValue of its response is 0x00000503");

                #endregion
            }
        }

        #endregion

        #region Common methods

        /// <summary>
        /// Get Opened Stream Handle.
        /// </summary>
        /// <param name="messageHandle">The message handle</param>
        /// <returns>Return the opened stream handle</returns>
        private uint GetOpenedStreamHandle(uint messageHandle)
        {
            RopOpenStreamRequest openStreamRequest;
            RopOpenStreamResponse openStreamResponse;

            PropertyTag tag;
            tag.PropertyId = this.propertyDictionary[PropertyNames.UserSpecified].PropertyId;
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
        /// <param name="streamHandle">The opened stream handle</param>
        private void WriteStream(uint streamHandle)
        {
            RopWriteStreamRequest writeStreamRequest;
            RopWriteStreamResponse writeStreamResponse;

            byte[] data = Encoding.ASCII.GetBytes(SampleStreamData + "\0");
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