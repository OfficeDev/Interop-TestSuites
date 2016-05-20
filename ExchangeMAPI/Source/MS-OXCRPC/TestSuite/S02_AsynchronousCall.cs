namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that refer to methods on AsyncEMSMDB interface
    /// </summary>
    [TestClass]
    public class S02_AsynchronousCall : TestSuiteBase
    {
        #region Variable
        /// <summary>
        /// Indicates whether events are pending for the client on the Session Context on the server.
        /// </summary>
        private bool isNotificationPending;

        /// <summary>
        /// A Boolean indicates whether the server disable asynchronous RPC notification.
        /// </summary>
        private bool isDisableAsyncRPCNotification = false;

        /// <summary>
        /// Declares a delegate for a method that returns a uint.
        /// </summary>
        /// <returns>Returns a delegate object.</returns>
        private delegate uint MethodCaller();
        #endregion

        #region Test Class Initialization and Cleanup
        /// <summary>
        /// Initializes the test class before running the test cases in the class.
        /// </summary>
        /// <param name="context">Context of test class</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class after running the test cases in the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This case tests methods on the AsyncEMSMDB interface without pending events.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S02_TC01_TestWithoutPendingEvent()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed, and send Session Context Handle (CXH) to EcDoAsyncConnectEx for testing EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoAsyncConnectEx
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R713, the returned value of method EcDoAsyncConnectEx is {0}.", this.returnValue);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R713 
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                713,
                @"[In EcDoAsyncConnectEx Method (opnum 14)] Return Values: If the method succeeds, the return value is 0.");

            #endregion

            #region Call EcDoAsyncWaitEx
            DateTime startTime = DateTime.Now;
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out this.isNotificationPending);
            DateTime endTime = DateTime.Now;
            TimeSpan interval = endTime.Subtract(startTime);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R704, the returned value of method EcDoAsyncWaitEx is {0}.", this.returnValue);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R704
            // Method EcDoAsyncWaitEx succeed indicates that method EcDoAsyncConnectEx binds a session context handle can be used in calls to EcDoAsyncWaitEx.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                704,
                @"[In EcDoAsyncConnectEx Method (opnum 14)] The EcDoAsyncConnectEx method binds a session context handle returned from the EcDoConnectEx method, as specified in section 3.1.4.1, to a new asynchronous context handle that can be used in calls to the EcDoAsyncWaitEx method in the AsyncEMSMDB interface, as specified in section 3.3.4.1.");

            if (Common.IsRequirementEnabled(1930, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1930, there are {0} pending events for the client on the Session Context on the server.The EcDoAsyncWaitEx execute time is {1}", this.isNotificationPending ? string.Empty : "not", interval.TotalMinutes);

                // Because above step not trigger any event, so if isNotificationPending is false and the execute time of EcDoAsyncWaitEx larger than 5 minutes.R1930 will be verified.
                bool isR1930Verified = this.isNotificationPending == false && interval.TotalSeconds >= 290;

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1930
                Site.CaptureRequirementIfIsTrue(
                    isR1930Verified,
                    1930,
                    @"[In Appendix B: Product Behavior] Implementation does return the call and will not set the NotificationPending flag in the pulFlagsOut field, If no events are available within five minutes of the time that the client last accessed the server through a call to EcDoRpcExt2. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1907, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1907, there are {0} pending events for the client on the Session Context on the server. The EcDoAsyncWaitEx execute time is {1}", this.isNotificationPending ? string.Empty : "not", interval.TotalMinutes);

                bool isR1907Verified = interval.TotalSeconds >= 290 && interval.TotalSeconds < 360;

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1907
                Site.CaptureRequirementIfIsTrue(
                    isR1907Verified,
                    1907,
                    @"[In Appendix B: Product Behavior] Implementation does complete the call every 5 minutes regardless of the client's last activity time. [In Appendix B: Product Behavior] <37> Section 3.3.4.1: Exchange 2007 completes the call every 5 minutes regardless of the client's last activity time.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1338");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1338
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                1338,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] Return Values: If the method succeeds, the return value is 0.");

            if (Common.IsRequirementEnabled(1922, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1922");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1922
                // If this method is invoked successfully, it means EcDoAsyncWaitEx method is supported. The return value 0 means success.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1922,
                    @"[In Appendix B: Product Behavior] <36> Section 3.3.4: Implementation does support AsyncEMSMDB method EcDoAsyncWaitEx. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }

            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            #endregion
        }

        /// <summary>
        /// This case tests methods on the AsyncEMSMDB interface with pending events.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S02_TC02_TestWithPendingEvent()
        {
            this.CheckTransport();

            #region Initializes Server and Client
            this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, this.authenticationService, this.userName, this.password);
            Site.Assert.IsTrue(this.returnStatus, "The returned status is {0}. TRUE means that initializing the server and client to call EcDoAsyncWaitEx successfully, and FALSE means that initializing the server and client to call EcDoAsyncWaitEx failed.", this.returnStatus);
            #endregion

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed and send Session Context Handle (CXH) to EcDoRpcExt2 for testing EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 method with RopLogon as rgbIn
            // Parameter inObjHandle is no use for RopLogon command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region Call EcDoAsyncConnectEx
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoAsyncConnectEx should succeed and sends ACXH to EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Register events on server
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopRegisterNotification, this.objHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopRegisterNotificationResponse registerNotificationResponse = (RopRegisterNotificationResponse)this.response;
            Site.Assert.AreEqual<uint>(0, registerNotificationResponse.ReturnValue, "RopRegisterNotification should succeed and 0 is expected to be returned. The returned value is {0}.", registerNotificationResponse.ReturnValue);
            #endregion

            #region Call EcDoAsyncWaitEx
            // Trigger the event
            bool isCreateMailSuccess = this.oxcrpcControlAdapter.CreateMailItem();
            Site.Assert.IsTrue(isCreateMailSuccess, "CreateMailItem method should execute successfully.");

            this.returnValue = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out this.isNotificationPending);
            Site.Assert.AreEqual<uint>(0, this.returnValue, @"EcDoAsyncWaitEx should succeed to check whether the NotificationPending flag is set in the pulFlagsOut field, on AsyncEMSMDB method if an event is pending. '0' is expected to be returned. The returned value is {0}.", this.returnValue);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1231, there are {0} pending events for the client on the Session Context on the server.", this.isNotificationPending ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1231
            Site.CaptureRequirementIfIsTrue(
                this.isNotificationPending,
                1231,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] If an event is pending, the server completes the call immediately and returns the NotificationPending flag in the pulFlagsOut parameter.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R19");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R19
            // Because client can use the asynchronous context handle to call EcDoAsyncWaitEx method successful.
            // So R19 will be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                19,
                @"[In ACXH Data Type] The AXCH data type is an asynchronous context handle to be used with an AsyncEMSMDB interface, as specified in section 3.3 and section 3.4.");

            #endregion

            #region Call EcDoRpcExt2 with no ROP in rgbIn to get the notify information
            // Parameter inObjHandle and auxInfo are no use for null ROP command, so set them to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.WithoutRops, this.unusedInfo, this.unusedInfo);
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pcbOut = ConstValues.ValidpcbOut;
            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, @"EcDoRpcExt2 should succeed to get the RopNotifyResponse. '0' is expected to be returned. The returned value is {0}.", this.returnValue);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1243, the ROP response is {0}", (RopNotifyResponse)this.response);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1243
            // According to the Open Specification MS-OXCNOTIF, the event details are in the RopNotifyResponse. If the rgbOut can be converted (parsed) to RopNotifyResponse, this requirement will be verified.
            Site.CaptureRequirementIfIsNotNull(
                (RopNotifyResponse)this.response,
                1243,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] [pulFlagsOut] [Flag NotificationPending] The server will return the event details in the ROP response buffer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1229, the ROP response is {0}", (RopNotifyResponse)this.response);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1229
            // If the response of method EcDoRpcExt2 is not null, it indicates that method EcDoAsyncWaitEx have been completed because server will return the event details in the ROP response buffer.
            Site.CaptureRequirementIfIsNotNull(
                (RopNotifyResponse)this.response,
                1229,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] The EcDoAsyncWaitEx method is an asynchronous call that the server does not complete until events are pending on the Session Context, up to a 5-minute duration of no client activity.");

            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and CXH used for testing EcDoAsyncWaitEx should be released. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case tests calling EcDoAsyncWaitEx when asynchronous context handle becomes invalid.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S02_TC03_TestInvalidAsynchronousContextHandle()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed, and send Session Context Handle (CXH) to EcDoAsyncConnectEx for testing EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 method with RopLogon as rgbIn
            // Parameter inObjHandle is no use for RopLogon command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region Call EcDoAsyncConnectEx
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoAsyncConnectEx should succeed and sends ACXH to EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Register events on server
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopRegisterNotification, this.objHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopRegisterNotificationResponse registerNotificationResponse = (RopRegisterNotificationResponse)this.response;
            Site.Assert.AreEqual<uint>(0, registerNotificationResponse.ReturnValue, "RopRegisterNotification should succeed and '0' is expected to be returned. The returned value is {0}.", registerNotificationResponse.ReturnValue);
            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and CXH used for testing EcDoAsyncWaitEx should be released. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoAsyncWaitEx
            uint returnValueForEcDoAsyncWaitEx = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out this.isNotificationPending);
            #endregion

            #region Call EcDoRpcExt2 with no ROP in rgbIn to get the notify information
            // Parameter inObjHandle and auxInfo are no use for null ROP command, so set them to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.WithoutRops, this.unusedInfo, this.unusedInfo);
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pcbOut = ConstValues.ValidpcbOut;
            uint returnValueForEcDoRpcExt2 = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            #region Capture code.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1213, server returns {0} when call EcDoAsyncWaitEx, server returns {1} when call EcDoRpcExt2.", returnValueForEcDoAsyncWaitEx, returnValueForEcDoRpcExt2);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1213
            bool isVerifiedR1213 = returnValueForEcDoAsyncWaitEx != 0 || returnValueForEcDoRpcExt2 != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1213,
                1213,
                @"[In Abstract Data Model] When the session context is destroyed, the asynchronous context handle becomes invalid and will be rejected if used.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1269,server returns {0} when call EcDoAsyncWaitEx, server returns {1} when call EcDoRpcExt2.", returnValueForEcDoAsyncWaitEx, returnValueForEcDoRpcExt2);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1269
            bool isVerifiedR1269 = returnValueForEcDoAsyncWaitEx != 0 || returnValueForEcDoRpcExt2 != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1269,
                1269,
                @"[In Abstract Data Model] When the client connection is lost, the asynchronous context handle becomes invalid and will be rejected if used.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This case tests calling EcDoAsyncWaitEx unsuccessfully.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S02_TC04_TestEcDoAsyncWaitExFail()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed, and send Session Context Handle (CXH) to EcDoAsyncConnectEx for testing EcDoAsyncWaitEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoAsyncConnectEx
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R713, the returned value of method EcDoAsyncConnectEx is {0}.", this.returnValue);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R713 
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                713,
                @"[In EcDoAsyncConnectEx Method (opnum 14)] Return Values: If the method succeeds, the return value is 0.");

            #endregion

            #region Call EcDoAsyncWaitEx when pacxh is invalid
            this.pcxhInvalid = (IntPtr)ConstValues.InvalidPcxh;
            this.returnValueForInvalidCXH = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pcxhInvalid, out this.isNotificationPending);
            Site.Assert.AreNotEqual<uint>(0, this.returnValueForInvalidCXH, "EcDoAsyncWaitEx should not succeed if the CXH is invalid. '0' isn't expected to be returned. The returned value is {0}.", this.returnValueForInvalidCXH);
            #endregion

            #region Capture code.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1228, the return value of EcDoAsyncWaitEx with ACXH valid is {0}, the return value of EcDoAsyncWaitEx with ACXH invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1228
            // When the ACXH is returned from EcDoAsyncWaitEx method, the call is successful. While the ACXH is not returned from the EcDoAsyncWaitEx method, the call is failed.
            // If the code can reach here, this requirement is verified.
            bool isVerifyR1228 = (this.returnValue == ResultSuccess) && (this.returnValueForInvalidCXH != ResultSuccess);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1228,
                1228,
                @"[In Message Processing Events and Sequencing Rules] Method EcDoAsyncWaitEx: The method requires an active asynchronous context handle returned from the EcDoAsyncConnectEx method on the EMSMDB interface.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1232, the return value of EcDoAsyncWaitEx with ACXH valid is {0}, the return value of EcDoAsyncWaitEx with ACXH invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1232
            // Since the context of R1232 is similar with R1228, if the R1228 is verified, it means R1232 is verified.
            bool isVerifyR1232 = isVerifyR1228;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1232,
                1232,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] This call [EcDoAsyncWaitEx] requires an active asynchronous context handle to be returned from the EcDoAsyncConnectEx method on the EMSMDB interface, as specified in section 3.1.4.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R310, the return value of EcDoAsyncWaitEx with ACXH valid is {0}, the return value of EcDoAsyncWaitEx with ACXH invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R310
            // Since the context of R310 is similar with R1228, if the R1228 is verified, it means R310 is verified.
            bool isVerifyR310 = isVerifyR1228;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR310,
                310,
                @"[In Protocol Details] All method calls that require a valid asynchronous context handle [EcDoAsyncWaitEx] are listed in the following table.");

            if (Common.IsRequirementEnabled(1908, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1908");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1908
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    TestSuiteBase.ResultSuccess,
                    this.returnValueForInvalidCXH,
                    1908,
                    @"[In Appendix B: Product Behavior] Implementation does reject the request if the asynchronous context handle is invalid. (<38> Section 3.3.4.1: Exchange 2007 and Exchange 2010 follow this behavior.)");
            }
            #endregion

            #region Call EcDoAsyncWaitEx with an EcDoAsyncWaitEx call outstanding on this ACXH
            uint returnValueOfFirstCall = 0;
            uint returnValueOfSecondCall = 0;

            MethodCaller asyncThread = new MethodCaller(
                () =>
                {
                    return this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out this.isNotificationPending);
                });

            IAsyncResult result = asyncThread.BeginInvoke(null, null);

            returnValueOfSecondCall = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out this.isNotificationPending);
            returnValueOfFirstCall = asyncThread.EndInvoke(result);
            Site.Log.Add(LogEntryKind.Debug, string.Format("The return value of the first EcDoAsyncWaitEx method is {0}", returnValueOfFirstCall));
            Site.Log.Add(LogEntryKind.Debug, string.Format("The return value of the second EcDoAsyncWaitEx method is {0}", returnValueOfSecondCall));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1343");

            bool isR1343Verified = returnValueOfSecondCall == 0x000007EE || returnValueOfFirstCall == 0x000007EE;

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1343
            Site.CaptureRequirementIfIsTrue(
                isR1343Verified,
                1343,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] [Return Values] [Rejected (0x000007EE)] An EcDoAsyncWaitEx method call is already outstanding on this asynchronous context handle.<38>");

            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            #endregion
        }

        /// <summary>
        /// This case tests calling EcDoAsyncConnectEx when server has disabled the asynchronous RPC notifications.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S02_TC05_TestEcDoAsyncConnectExWithDisableAsynchronous()
        {
            this.CheckTransport();

            #region Call DisableAsyncRPCNotification method to disable asynchronous RPC notifications.
            // Disable the asynchronous RPC notification on server
            this.oxcrpcControlAdapter.DisableAsyncRPCNotification();

            // On Exchange 2013, when test case restart the service,
            // the service is not ready immediately for communicate with server over RPC.
            // So test case need call EcDoConnectEx and EcDoRpcExt2 to check the service whether is ready for communicate with server over RPC.
            bool isRestartSuccess = this.CheckServiceStartSuccess();
            Site.Assert.IsTrue(isRestartSuccess, "The service should start success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Asynchronous RPC notification is disabled on server through SUT control Adapter.");
            this.isDisableAsyncRPCNotification = true;
            #endregion

            #region Initializes Server and Client
            this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, this.authenticationService, this.userName, this.password);
            Site.Assert.IsTrue(this.returnStatus, "The returned status is {0}. TRUE means that initializing the server and client in order to call the following EcDoAsyncConncet successfully, and FALSE means that initializing the server and client in order to call the following EcDoAsyncConncet failed.", this.returnStatus);
            #endregion

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed, and send Session Context Handle (CXH) to EcDoAsyncConnectEx. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Bind a CXH to an ACXH
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);

            if (Common.IsRequirementEnabled(1812, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1812");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1812
                Site.CaptureRequirementIfAreEqual<uint>(
                0x000007EE,
                this.returnValue,
                1812,
                @"[In Appendix B: Product Behavior] Implementation does return the ecRejected error code, when the Server has asynchronous RPC notifications disabled. (Microsoft Exchange Server 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1942, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1942");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1942
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x000007EE,
                    this.returnValue,
                    1942,
                    @"[In Appendix B: Product Behavior] Implementation does return ecRejected (0x000007EE) when client either polls for notifications or calls EcRRegisterPushNotifications. (Microsoft Exchange Server 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1941, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1941");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1941
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x000007EE,
                    this.returnValue,
                    1941,
                    @"[In Appendix B: Product Behavior] Implementation does not return [ecRejected (0x000007EE)] when Client either polls for notifications or calls EcRRegisterPushNotifications. [In Appendix B: Product Behavior]  <26> Section 3.1.4.4: Exchange 2010, Exchange 2013, and Exchange 2016 do not return the ecRejected error code.");
            }

            if (Common.IsRequirementEnabled(1757, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1757");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1757
                Site.CaptureRequirementIfAreNotEqual<uint>(
                0x000007EE,
                this.returnValue,
                1757,
                @"[In Appendix B: Product Behavior] Implementation does not return the ecRejected error code, when the Server has asynchronous RPC notifications disabled. (<26> Section 3.1.4.4: Exchange 2010, Exchange 2013, and Exchange 2016 do not return the ecRejected error code.)");
            }

            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect with the server should succeed and CXH for testing EcDoAsyncConnectEx should be released. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EnableAsyncRPCNotification method to enable asynchronous RPC notifications.
            // Enable the asynchronous RPC notification on server
            this.oxcrpcControlAdapter.EnableAsyncRPCNotification();

            // On Exchange 2013, when test case restart the service,
            // the service is not ready immediately for connect server over RPC.
            // So test case need call EcDoConnectEx to check the service whether is ready for connect server over RPC.
            isRestartSuccess = this.CheckServiceStartSuccess();
            Site.Assert.IsTrue(isRestartSuccess, "The service should start success.");
            this.isDisableAsyncRPCNotification = false;
            #endregion
        }

        /// <summary>
        /// Clean up the test case after running it
        /// </summary>
        protected override void TestCleanup()
        {
            #region Call EnableAsyncRPCNotification method to enable asynchronous RPC notifications.
            if (this.isDisableAsyncRPCNotification == true)
            {
                // Enable the asynchronous RPC notification on server
                this.oxcrpcControlAdapter.EnableAsyncRPCNotification();
                this.isDisableAsyncRPCNotification = false;
            }
            #endregion

            base.TestCleanup();
        }

        /// <summary>
        /// Check whether the service started successfully.
        /// </summary>
        /// <returns>A Boolean value indicates whether service started successfully.</returns>
        private bool CheckServiceStartSuccess()
        {
            int tryToConnectCount = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("ConnectRetryCount", Site));

            #region Initializes Server and Client
            this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, this.authenticationService, this.userName, this.password);
            Site.Assert.IsTrue(this.returnStatus, "The returned status is {0}. TRUE means that initializing the server and client in order to call the following EcDoAsyncConncet successfully, and FALSE means that initializing the server and client in order to call the following EcDoAsyncConncet failed.", this.returnStatus);
            #endregion

            #region Call EcDoConnectEx to try to connect with Server
            while (tryToConnectCount < retryCount)
            {
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                    ref this.pcxh,
                    TestSuiteBase.UlIcxrLinkForNoSessionLink,
                    ref this.pulTimeStamp,
                    null,
                    this.userDN,
                    ref this.pcbAuxOut,
                    this.rgwClientVersion,
                    out this.rgwBestVersion,
                    out this.picxr);

                if (this.returnValue == 0)
                {
                    this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
                    this.pcbOut = ConstValues.ValidpcbOut;
                    this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

                    this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                        ref this.pcxh,
                        PulFlags.NoCompression | PulFlags.NoXorMagic,
                        this.rgbIn,
                        ref this.pcbOut,
                        null,
                        ref this.pcbAuxOut,
                        out this.response,
                        ref this.responseSOHTable);

                    if (this.returnValue == 0)
                    {
                        this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                        return true;
                    }
                    else
                    {
                        this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                    }
                }

                tryToConnectCount++;
                Thread.Sleep(waitTime*3);
            }

            return false;
            #endregion
        }
    }
}