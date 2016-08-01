namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Net.Sockets;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that refer to methods on EMSMDB interface 
    /// </summary>
    [TestClass]
    public class S01_SynchronousCall : TestSuiteBase
    {
        #region Variable
        /// <summary>
        /// An integer indicates a value for reserved field. This value can be any 4-bytes value.
        /// </summary>
        private const int ReserveDefault = 0;

        /// <summary>
        /// Contains an auxiliary payload buffer
        /// </summary>
        private byte[] rgbAuxIn;

        /// <summary>
        /// The subfolder ID list that be created.
        /// </summary>
        private List<ulong> subFolderIds = new List<ulong>();
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
       
        #region Test cases
        /// <summary>
        /// This test case verifies the session context related to method EcDoConnectEx and 
        /// whether methods EcDoDisconnect and EcDoRpcExt2 require a valid CXH.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()] 
        public void MSOXCRPC_S01_TC01_TestEcDoConnectEx()
        {
            this.CheckTransport();

            #region Call EcDoConnectEx method to establish a new Session Context with the server.
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

            uint firstPulTimeStamp = this.pulTimeStamp;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R582");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R582
            Site.CaptureRequirementIfAreEqual<long>(
                0,
                this.returnValue,
                582,
                @"[In EcDoConnectEx Method (Opnum 10)] Return Values: If the method succeeds, the return value is 0.");

            #endregion

            #region Client connects with server secondly using session context linking
            IntPtr secondCxh = IntPtr.Zero;
            ushort secondPicxr;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            if (Common.IsRequirementEnabled(508, this.Site))
            {
                this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                    ref secondCxh,
                    this.picxr,
                    ref this.pulTimeStamp,
                    null,
                    this.userDN,
                    ref this.pcbAuxOut,
                    this.rgwClientVersion,
                    out this.rgwBestVersion,
                    out secondPicxr);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx for testing session context linking should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            else
            {
                string normalUserDN = Common.GetConfigurationPropertyValue("NormalUserEssdn", Site);
                this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                    ref secondCxh,
                    this.picxr,
                    ref this.pulTimeStamp,
                    null,
                    normalUserDN,
                    ref this.pcbAuxOut,
                    this.rgwClientVersion,
                    out this.rgwBestVersion,
                    out secondPicxr);

                if (Common.IsRequirementEnabled(1850, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1850");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1850
                    // Because a Session Context has been created on server during the above step.
                    // When calling EcDoConnectEx method to link this Session Context, this session must be found on server 
                    // and server will link the Session Context created by this call with the one found.
                    // So if the EcDoConnectEx method returns success then R1850 will be verified.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        0,
                        this.returnValue,
                        1850,
                        @"[In Appendix B: Product Behavior] Implementation link the Session Context created by this call with the one found, If a session is found. (Microsoft Exchange Server 2007 follows this behavior.)");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R557");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R557
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    557,
                    @"[In EcDoConnectEx Method (Opnum 10)] [pulTimeStamp] If so [If the server supports Session Context linking, the server verifies that there is a Session Context state with the unique identifier ulIcxrLink and the server has a creation time stamp equal to the value passed in this parameter], the server MUST link the Session Context created by this [method EcDoConnectEx] call with the one found.");
            }

            if (Common.IsRequirementEnabled(1943, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1943");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1943
                Site.CaptureRequirementIfAreNotEqual<ushort>(
                    this.picxr,
                    secondPicxr,
                    1943,
                    @"[In Appendix B: Product Behavior] Implementation does not assign two active Session Contexts the same session index value. (Microsoft Exchange Server 2007 follows this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4790");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R4790
            Site.CaptureRequirementIfAreNotEqual<IntPtr>(
                secondCxh,
                this.pcxh,
                4790,
                @"[In EcDoConnectEx Method (Opnum 10)] pcxh: On success, the server MUST return a value to be used as a session context handle, and the value is not the same as the next EcDoConnectEx successful call.");
            #endregion

            #region Capture code.
            if (!Common.IsRequirementEnabled(508, this.Site))
            {
                #region Logon on using the first CXH
                // Parameter inObjHandle is not used for RopLogon, so set it to unUsedInfo.
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
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                RopLogonResponse logonResponse = (RopLogonResponse)this.response;
                Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
                this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
                #endregion

                #region Register events on server using the first CXH
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
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                RopRegisterNotificationResponse registerNotificationResponse = (RopRegisterNotificationResponse)this.response;
                Site.Assert.AreEqual<uint>(0, registerNotificationResponse.ReturnValue, "RopRegisterNotification should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", registerNotificationResponse.ReturnValue);
                #endregion

                #region Logon using the second CXH
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogonNormalUser, this.unusedInfo, this.userPrivilege | (ulong)OpenFlags.UseAdminPrivilege);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref secondCxh,
                    PulFlags.NoCompression | PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed by using second CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                RopLogonResponse secondlogonResponse = (RopLogonResponse)this.response;
                Site.Assert.AreEqual<uint>(0, secondlogonResponse.ReturnValue, "RopLogon should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", secondlogonResponse.ReturnValue);
                this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][secondlogonResponse.OutputHandleIndex];
                #endregion

                #region Register events on server using the second CXH
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopRegisterNotification, this.objHandle, secondlogonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref secondCxh,
                    PulFlags.NoCompression | PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed by using second CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                registerNotificationResponse = (RopRegisterNotificationResponse)this.response;
                Site.Assert.AreEqual<uint>(0, registerNotificationResponse.ReturnValue, "RopRegisterNotification should succeed by using the first CXH and '0' is expected to be returned. The returned value is {0}.", registerNotificationResponse.ReturnValue);
                #endregion

                // Trigger event to the first session context
                bool isCreateMailSuccess = this.oxcrpcControlAdapter.CreateMailItem();
                Site.Assert.IsTrue(isCreateMailSuccess, "CreateMailItem method should execute successfully.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Sends a mail to the specified mailbox successfully through SUT control Adapter.");

                int count = 0;
                int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
                int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
                while (true)
                {
                    #region Get the RopPending information using the second CXH
                    // Parameter inObjHandle and auxInfo are not used for null ROP command, so set them to unUsedInfo.
                    this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.WithoutRops, this.unusedInfo, this.unusedInfo);
                    this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                    this.pcbOut = ConstValues.ValidpcbOut;
                    this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                        ref secondCxh,
                        PulFlags.NoCompression | PulFlags.NoXorMagic,
                        this.rgbIn,
                        ref this.pcbOut,
                        null,
                        ref this.pcbAuxOut,
                        out this.response,
                        ref this.responseSOHTable);
                    #endregion

                    if (this.response == null)
                    {
                        count++;
                        if (count > retryCount)
                        {
                            break;
                        }

                        // Wait for the event to take effect
                        Thread.Sleep(waitTime);
                    }
                    else
                    {
                        break;
                    }
                }

                Site.Assert.IsNotNull(this.response, "The ROP response should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1477");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1477
                // The First CXH and Second CXH are linked, here the first CXH has events triggered, if the second CXH can get the pending event info of the first CXH,
                // this requirement is verified. The variable piCxr stores the sessionIndex of the First CXH.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    this.picxr,
                    ((RopPendingResponse)this.response).SessionIndex,
                    1477,
                    @"[In EcDoConnectEx Method (Opnum 10)] [piCxr] If Session Contexts are linked, a RopPending can be returned for the linked Session Context.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1177");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1177
                Site.CaptureRequirementIfAreEqual<ushort>(
                    this.picxr,
                    ((RopPendingResponse)this.response).SessionIndex,
                    1177,
                    @"[In Sending the EcDoConnectEx Method] [piCxr] If a client links Session Contexts, a RopPending ROP can be returned for any linked Session Context.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1175, the actual value is {0}", ((RopPendingResponse)this.response).SessionIndex);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1175
                // If any of the linked sessionIndex is returned, then R1175 will be verified.
                Site.CaptureRequirementIfIsNotNull(
                    ((RopPendingResponse)this.response).SessionIndex,
                    1175,
                    @"[In Sending the EcDoConnectEx Method] [piCxr] It is the session index returned in a RopPending ROP response ([MS-OXCROPS] section 2.2.14.3) on calls to the EcDoRpcExt2 method, as specified in section 3.1.4.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R536, the actual value is {0}", ((RopPendingResponse)this.response).SessionIndex);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R536
                // If any of the linked sessionIndex is returned, this requirement is verified.
                bool isVerifyR536 = (
                    ((RopPendingResponse)this.response).SessionIndex == this.picxr) ||
                    (((RopPendingResponse)this.response).SessionIndex == secondPicxr);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR536,
                    536,
                    @"[In EcDoConnectEx Method (Opnum 10)] [piCxr] The server MUST also use the session index when returning a RopPending ROP response ([MS-OXCROPS] section 2.2.14.3) on calls to the EcDoRpcExt2 method, as specified in section 3.1.4.2,  to tell the client which Session Context has pending notifications.");
            }
            else
            {
                // Client must disconnect with server to release the second CXH when server does not support session context linking.
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref secondCxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed to release the second CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            #region Client connects with server thirdly using an invalid session context linking
            IntPtr thirdCxh = IntPtr.Zero;
            ushort thirdPicxr;
            uint thirdPulTimeStamp = 0;

            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref thirdCxh,
                ConstValues.Invalidpicxr,
                ref thirdPulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out thirdPicxr);

            if (Common.IsRequirementEnabled(1944, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1944");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1944
                // In the third call to EcDoConnectEx, the piCxr parameter is invalid, so the server will not find any Session Context to link.
                // If the server just returns 0 which means the connection is successful, then this requirement is verified.
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1944,
                    @"[In Appendix B: Product Behavior] If no such Session Context state is found, the Implementation does not fail the EcDoConnectEx call, but simply does not do linking.  (Microsoft Exchange Server 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(508, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R508");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R508
                Site.CaptureRequirementIfAreEqual<uint>(
                    firstPulTimeStamp,
                    this.pulTimeStamp,
                    508,
                    @"[In Appendix B: Product Behavior] Implementation does return the same value in the pulTimeStamp that was passed in. [In Appendix B: Product Behavior] [<8> Section 3.1.4.1] [In Exchange 2010, Exchange 2013 and Exchange 2016, if ulIcxrLink is not 0xFFFFFFFF, then the server will not attempt to search for a session with the same Session Context and link to them,] it [method EcDoConnectEx] will then return the same value in the pulTimeStamp that was passed in.");
            }

            if (Common.IsRequirementEnabled(1435, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1435, the firstPulTimeStamp is: {0}, the pulTimeStamp is: {1}", firstPulTimeStamp, this.pulTimeStamp);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1435
                Site.CaptureRequirementIfAreEqual<uint>(
                    firstPulTimeStamp,
                    this.pulTimeStamp,
                    1435,
                    @"[In Appendix B: Product Behavior] Implementation does return the same value in the pulTimeStamp that was passed in. [In Appendix B: Product Behavior] [<11> Section 3.1.4.1] Rather [in Exchange 2010, Exchange 2013, and Exchange 2016, if ulIcxrLink is not 0xFFFFFFFF, then the server will not attempt to search for a session with the same Session Context and link to the server], it [the server] will then return the same value in the pulTimeStamp that was passed in.");
            }
            #endregion 

            #region Client disconnects with server to release the third CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref thirdCxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed to release the third CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoDisconnect with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R408");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R408
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                408,
                @"[In EcDoDisconnect Method (opnum 1)] Error Codes: If the method succeeds, the return value is 0.");
            #endregion

            #region Call EcDoDisconnect when pcxh is invalid.
            this.pcxhInvalid = (IntPtr)ConstValues.InvalidPcxh;
            this.returnValueForInvalidCXH = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxhInvalid);
            Site.Assert.AreNotEqual<uint>(0, this.returnValueForInvalidCXH, "EcDoDisconnect should not succeed by using an invalid CXH and '0' isn't expected to be returned. The returned value is {0}.", this.returnValueForInvalidCXH);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R404, the return value of calling EcDoDisconnect with pcxh valid is {0}, the return value of calling EcDoDisconnect with pcxh invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R404
            // When the CXH is returned from EcDoConnectEx method, the call is successful. While the CXH is not returned from the EcDoConnectEx method, the call is failed.
            // If the code can reach here, this requirement is verified.
            bool isVerifyR404 = (this.returnValue == ResultSuccess) && (this.returnValueForInvalidCXH != ResultSuccess);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR404,
                404,
                @"[In EcDoDisconnect Method (opnum 1)] This [the method EcDoDisconnect] call requires an active session context handle from the EcDoConnectEx method, as specified in section 3.1.4.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R372, the return value of calling EcDoDisconnect with pcxh valid is {0}, the return value of calling EcDoDisconnect with pcxh invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R372
            // Since the context of R372 is similar with R404, if the R404 is verified, it means R372 is verified.
            bool isVerifyR372 = isVerifyR404;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR372,
                372,
                @"[In Message Processing Events and Sequencing Rules] [EcDoDisconnect] The method requires an active session context handle to be returned from the EcDoConnectEx method, as specified in section 3.1.4.1.");

            #endregion

            #region Call EcDoRpcExt2 with a valid CXH after destroying the valid CXH.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R401");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R401
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.returnValue,
                401,
                @"[In EcDoDisconnect Method (opnum 1)] The EcDoDisconnect method closes a Session Context with the server.");
            #endregion
        }

        /// <summary>
        /// This test case verifies invalid parameters in EcDoConnectEx method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC02_TestInvalidParameterForMethodEcDoConnectEx()
        {
            this.CheckTransport();

            byte[] payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(ReserveDefault);

            #region Client connects with Server when the client version is invalid
            // 0x1, 0x2, 0x3 are older values for rgwClientVersion than that required by the server.
            ushort[] rgwClientVersionTemp = new ushort[3] { 0x1, 0x2, 0x3 };
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                rgwClientVersionTemp,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreNotEqual<uint>(0, this.returnValue, "EcDoConnectEx should fail when the client and server versions are not compatible and '0' isn't expected to be returned. The returned value is {0}.", this.returnValue);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R483");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R483
            Site.CaptureRequirementIfAreEqual<IntPtr>(
                IntPtr.Zero,
                this.pcxh,
                483,
                @"[In EcDoConnectEx Method (Opnum 10)] [pcxh] On failure, the server MUST return a zero value as the session context handle.");

            #endregion

            #region Client connects with Server when cbAuxIn is equal to TooBigcbAuxIn
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.rgbAuxIn = new byte[ConstValues.TooBigcbAuxIn];
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);

            if (Common.IsRequirementEnabled(4875, this.Site))
            {            
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4875");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R4875
                // The condition that cbAuxIn is larger than 0x00001008 bytes is controlled by "TooBigcbAuxIn" because the value of "TooBigcbAuxIn" is defined as a constant with value "4105(0x1009)".
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000006F7,
                    this.returnValue,
                    4875,
                    @"[In Appendix B: Product Behavior] Implementation does fail with the RPC status code RPC_X_BAD_STUB_DATA (0x000006F7) if the value of cbAuxIn is larger than 0x00001008 bytes in size. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Client connects with Server when cbAuxIn is equal to TooSmallcbAuxIn
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
            this.picxr = 0;
            this.rgbAuxIn = new byte[ConstValues.TooSmallcbAuxIn];
            Array.Copy(payload, this.rgbAuxIn, ConstValues.TooSmallcbAuxIn);
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);

            if (Common.IsRequirementEnabled(1436, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1436, the return value is {0}", this.returnValue);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1436
                // The condition that cbAuxIn is less than 0x00000008 is controlled by constant "TooSmallcbAuxIn" defined in ConstValues.cs.
                int lengthOfcbAuxIn = this.rgbAuxIn.Length;
                bool isVerifyR1436 = (0x00000000 == this.returnValue) && (lengthOfcbAuxIn > 0x00000000 && lengthOfcbAuxIn < 0x00000008);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1436,
                    1436,
                    @"[In Appendix B: Product Behavior] Implementation does not fail if cbAuxIn is greater than 0x00000000 and less than 0x00000008. <14> Section 3.1.4.1: Exchange 2007 does not fail if cbAuxIn is greater than 0x00000000 and less than 0x00000008.");
            }

            if (Common.IsRequirementEnabled(1940, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1940, the return value is {0}", this.returnValue);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1940
                // The condition that cbAuxIn is less than 0x00000008 is controlled by constant "TooSmallcbAuxIn" defined in ConstValues.cs.
                int lengthOfcbAuxIn = this.rgbAuxIn.Length;
                bool isVerifyR1940 = (0x80040115 == this.returnValue) && (lengthOfcbAuxIn > 0x00000000 && lengthOfcbAuxIn < 0x00000008);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1940,
                    1940,
                    @"[In Appendix B: Product Behavior] Implementation does fail with ecRpcFailed (0x80040115) if this value is greater than 0x00000000 and less than 0x00000008. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }
            #endregion

            #region Client connects with Server when pcbAuxOut is equal to TooBigpcbAuxOut
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.TooBigpcbAuxOut;
            this.pulTimeStamp = 0;
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

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R579");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R579
            // The condition that pcbAuxOut is larger than 0x00001008 bytes is controlled by constant "TooBigpcbAuxOut" defined in ConstValues.cs.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000006F7,
                this.returnValue,
                579,
                @"[In EcDoConnectEx Method (Opnum 10)] [pcbAuxOut] If this value on input is larger than 0x00001008, the server MUST fail with the RPC status code RPC_X_BAD_STUB_DATA (0x000006F7).");

            #endregion
        }

        /// <summary>
        /// This test case verifies invalid values for parameter cbIn in EcDoRpcExt2 method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC03_TestInvalidParametercbInForEcDoRpcExt2()
        {
            this.CheckTransport();

            if (Common.IsRequirementEnabled(1508, this.Site))
            {
                #region Client connects with Server
                this.pcxh = IntPtr.Zero;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.pulTimeStamp = 0;
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
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                #endregion

                #region Call EcDoRpcExt2 when cbIn is equal to TooSmallcbIn
                this.rgbIn = new byte[ConstValues.TooSmallcbIn];
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1508");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1508
                // The condition that cbIn is less than 0x00000008 is controlled by constant "TooSmallcbIn" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040115,
                    this.returnValue,
                    1508,
                    @"[In Appendix B: Product Behavior] Implementation does fail with error code ecRpcFailed (0x80040115) if the request buffer is smaller than the size of RPC_HEADER_EXT (0x00000008 bytes). (Microsoft Exchange Server 2010 SP2 and above follow this behavior.)");
                #endregion

                #region Client disconnects with Server
                if (this.pcxh != IntPtr.Zero)
                {
                    this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                    Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                }
                #endregion

                // Wait one second before sending next invalid request to avoid many invalid requests received by server in short time.
                System.Threading.Thread.Sleep(1000);
            }

            #region Client connects with Server
            // Since method EcDoRpcExt2 may destroy pcxh for implementation-specific error especially when input parameters are invalid, connect to server every time before calling method EcDoRpcExt2 to make sure the error code is just caused by the invalid parameter.
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when cbIn is equal to BigcbIn
            this.rgbIn = new byte[ConstValues.BigcbIn];
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

            if (Common.IsRequirementEnabled(1374, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1374");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1374
                // According to the Error Codes of method EcDoRpcExt2, the value of ecRpcFormat is 0x000004b6.
                // The condition that cbIn is greater than 0x00008007 is controlled by constant "BigcbIn" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004b6,
                    this.returnValue,
                    1374,
                    @"[In Appendix B: Product Behavior] Implementation does fail with error code ecRpcFormat if the request buffer is larger than 0x00008007 bytes in size. <19> Section 3.1.4.2: Exchange 2007 and 2010 will fail with error code ecRpcFormat (0x000004B6) if the request buffer is larger than 0x00008007 bytes in size.");
            }

            if (Common.IsRequirementEnabled(2001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R2001");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R2001
                // The condition that cbIn is less than 0x00000008 is controlled by constant "TooSmallcbIn" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040115,
                    this.returnValue,
                    2001,
                    @"[In Appendix B: Product Behavior] Implementation does fail with error code ecRpcFailed (0x80040115) if the request buffer is larger than 0x00008007 bytes in size. (Microsoft Exchange Server 2010 Service Pack 2 (SP2), Microsoft Exchange Server 2013 Service Pack 1 (SP1), and Exchange 2016 will fail with error code ecRpcFailed (0x80040115) if the request buffer is larger than 0x00008007 bytes in size.)");
            }
            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            // Wait one second before sending next invalid request to avoid many invalid requests received by server in short time.
            System.Threading.Thread.Sleep(1000);

            if (Common.IsRequirementEnabled(1939, this.Site))
            {
                #region Client connects with Server
                // Since method EcDoRpcExt2 may destroy pcxh for implementation-specific error especially when input parameters are invalid, connect to server every time before calling method EcDoRpcExt2 to make sure the error code is just caused by the invalid parameter.
                this.pcxh = IntPtr.Zero;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.pulTimeStamp = 0;
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
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                #endregion

                #region Call EcDoRpcExt2 when cbIn is equal to TooBigcbIn
                this.rgbIn = new byte[ConstValues.TooBigcbIn];
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1939");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1939
                // According to the Error Codes of method EcDoRpcExt2, the value of RPC_X_BAD_STUB_DATA is 0x000006F7.
                // The condition that cbIn is greater than 0x00040000 is controlled by constant "TooBigcbIn" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000006F7,
                    this.returnValue,
                    1939,
                    @"[In Appendix B: Product Behavior] Implementation does fail with the RPC status code of RPC_X_BAD_STUB_DATA (0x000006F7) if the request buffer is larger than 0x00040000 bytes in size. (Microsoft Exchange Server 2010 Service Pack 2 (SP2) and above follow this behavior.)");
                #endregion

                #region Client disconnects with Server
                if (this.pcxh != IntPtr.Zero)
                {
                    this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                    Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                }
                #endregion
            }
        }

        /// <summary>
        /// This test case verifies invalid values for parameter pcbOut in EcDoRpcExt2 method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC04_TestInvalidParameterpcbOutForEcDoRpcExt2()
        {
            this.CheckTransport();
                        
            #region Client connects with Server
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when pcbOut is equal to SmallpcbOut
            // Parameter inObjHandle is no use for RopLogon command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.SmallpcbOut;
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

            if (Common.IsRequirementEnabled(1924, this.Site))
            {
                // Add the debug information 
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1924");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1924
                // The condition that pcbOut is less than 0x00008007 is controlled by constant "SmallpcbOut" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.returnValue,
                    1924,
                    @"[In Appendix B: Product Behavior] Implementation does fail with ecRpcFormat (0x000004B6) if the output buffer is less than 0x00008007. (Microsoft Exchange Server 2007 follows this behavior).");
            }            
            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            // Wait one second before sending next invalid request to avoid many invalid requests received by server in short time.
            System.Threading.Thread.Sleep(1000);

            #region Client connects with Server
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when pcbOut is equal to TooSmallpcbOut
            // Parameter inObjHandle is no use for RopLogon command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.TooSmallpcbOut;
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
            #endregion

            #region Capture code.
            if (Common.IsRequirementEnabled(664, this.Site))
            {
                // Add the debug information 
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R664");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R664
                // The condition that pcbOut is less than 0x00000008 is controlled by constant "TooSmallpcbOut" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040115,
                    this.returnValue,
                    664,
                    @"[In Appendix B: Product Behavior] Implementation does fail with error code ecRpcFailed (0x80040115) if the value in pcbOut on input is less than 0x00000008. (Microsoft Exchange Server 2010 follows this behavior).");
            }

			if (Common.IsRequirementEnabled(2002, this.Site))
            {
                // Add the debug information 
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R2002");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R2002
                // The condition that pcbOut is less than 0x00000008 is controlled by constant "TooSmallpcbOut" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000000,
                    this.returnValue,
                    2002,
                    @"[In Appendix B: Product Behavior] Implementation does succeed if output buffer is less than 0x00000008, but no request ROPs will have been processed. (Microsoft Exchange Server 2013 and Microsoft Exchange Server 2016 follow this behavior).");
            }
			
            if (Common.IsRequirementEnabled(1900, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1900");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1900
                // The condition that cbIn is less than 0x00000008 is controlled by constant "TooSmallpcbOut" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.returnValue,
                    1900,
                    @"[In Appendix B: Product Behavior] Implementation does fail with error code ecRpcFormat (0x000004B6) if the value in pcbOut is less than 0x00000008. (<20> Section 3.1.4.2: Exchange 2007, and Microsoft Exchange Server 2010 Service Pack 1 (SP1) fail with error code ecRpcFormat (0x000004B6) if the value in the cbIn parameter is less than 0x00000008. )");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R697");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R697
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.returnValue,
                    697,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [Return Values] [ecRpcFormat (0x000004B6)] The format of the request was found to be invalid.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R698");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R698
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.returnValue,
                    698,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [Return Values] [ecRpcFormat (0x000004B6)] This [ecRpcFormat] is a generic error that means the length was found to be invalid or the content was found to be invalid.");
            }
            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            // Wait one second before sending next invalid request to avoid many invalid requests received by server in short time.
            System.Threading.Thread.Sleep(1000);

            #region Client connects with Server
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when pcbOut is equal to TooBigpcbOut
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.TooBigpcbOut;
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

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R666");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R666
            // The condition that pcbOut is larger than 0x00040000 is controlled by constant "TooBigpcbOut" defined in ConstValues.cs.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000006F7,
                this.returnValue,
                666,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pcbOut] If the value in the pcbOut parameter on input is larger than 0x00040000, the server MUST fail with the RPC status code of RPC_X_BAD_STUB_DATA (0x000006F7).");

            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion
        }

        /// <summary>
        /// This test case verifies invalid values for parameter cbAuxIn in EcDoRpcExt2 method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC05_TestInvalidParametercbAuxInForEcDoRpcExt2()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcxh = IntPtr.Zero;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when cbAuxIn is equal to TooBigcbAuxIn
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.rgbAuxIn = new byte[ConstValues.TooBigcbAuxIn];

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                this.rgbAuxIn,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            if (Common.IsRequirementEnabled(4877, this.Site))
            {
                // Add the debug information 
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4877");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R4877
                // The condition that cbAuxIn is larger than 0x00001008 bytes is controlled by constant "TooBigcbAuxIn" defined in ConstValues.cs.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000006F7,
                    this.returnValue,
                    4877,
                    @"[In Appendix B: Product Behavior] Implementation does fail with return code 0x000006F7 if the request buffer value of the cbAuxIn parameter is larger than 0x00001008 bytes in size. (<22> Section 3.1.4.2: Exchange 2010 and above follow this behavior.)");
            }

            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            // Wait one second before sending next invalid request to avoid many invalid requests received by server in short time.
            System.Threading.Thread.Sleep(1000);

            if (Common.IsRequirementEnabled(1381, this.Site))
            {
                #region Client connects with Server
                this.pcxh = IntPtr.Zero;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.pulTimeStamp = 0;
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
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                #endregion

                #region Call EcDoRpcExt2 when cbAuxIn is equal to TooSmallcbAuxIn
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.rgbAuxIn = new byte[ConstValues.TooSmallcbAuxIn];

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref this.pcxh,
                    PulFlags.NoCompression | PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.pcbOut,
                    this.rgbAuxIn,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1381, the return value is {0}", this.returnValue);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1381
                // The condition that cbAuxIn is less than 0x00000008 is controlled by "TooSmallcbAuxIn" 
                // because the value of "TooSmallcbAuxIn" is defined as a constant with value "7(0x0007)".
                int lengthOfcbAuxIn = this.rgbAuxIn.Length;
                bool isVerifyR1381 = (0x80040115 == this.returnValue) && (lengthOfcbAuxIn > 0x00000000 && lengthOfcbAuxIn < 0x00000008);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1381,
                    1381,
                    @"[In Appendix B: Product Behavior] Implementation does fail with ecRpcFailed (0x80040115) if the cbAuxIn parameter is greater than 0x00000000 and less than 0x00000008. (<23> Section 3.1.4.2: Exchange 2010 follows this behavior.)");
            }
            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion
        }

        /// <summary>
        /// This test case verifies invalid values for parameter pcbAuxOut in EcDoRpcExt2 method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC06_TestInvalidParameterpcbAuxOutForEcDoRpcExt2()
        {
            this.CheckTransport();

            #region Client connects with Server

            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.pulTimeStamp = 0;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 when pcbAuxOut is equal to TooBigpcbAuxOut
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.TooBigpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R686");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R686
            // The condition that pcbAuxOut is larger than 0x00001008 bytes is controlled by constant "TooBigpcbAuxOut" defined in ConstValues.cs.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000006F7,
                this.returnValue,
                686,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pcbAuxOut] If this value on input is larger than 0x00001008, the server MUST fail with the RPC status code RPC_X_BAD_STUB_DATA (0x000006F7).");

            #endregion

            #region Client disconnects with Server
            if (this.pcxh != IntPtr.Zero)
            {
                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion
        }

        /// <summary>
        /// This test case mainly aims to test the requirements referring to the compression algorithm, 
        /// obfuscation algorithm and extended buffer packing for RopQueryRows. 
        /// It also covers the requirements marked as adapter for methods EcDoConnectEx, EcDoRpcExt2 and EcDoDisconnect.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC07_TestExtendedBuffer()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbOut = ConstValues.ValidpcbOut;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx is the precondition for EcDoRpcExt2 and should succeed. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Logon to Mailbox
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            // Since the server has already been configurated to return auxiliary buffers, 
            // pass the non-null rgbAuxIn on method EcDoRpcExt2 will make server return non-null rgbAuxOut.
            byte[] payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO_V2(ReserveDefault);
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);
            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                this.rgbAuxIn,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            uint logonHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region OpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, logonHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            uint payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and '0' is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openFolderResponse.OutputHandleIndex];
            #endregion

            // Need create 2500 subfolders for verifying requirements related to extended buffer packing for RopQueryRows.
            uint folderNumber = 2500;

            #region Create Subfolders
            for (int i = 0; i < folderNumber; i++)
            {
                // Create subFolder.
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateFolder, this.objHandle, (ulong)i);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();
                payloadCount = 0;

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref this.pcxh,
                    PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.rgbOut,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable,
                    out payloadCount,
                    ref this.rgbAuxOut);

                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
                RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.response;
                Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "RopCreateFolder should succeed and '0' is expected to be returned. The returned value is {0}.", createFolderResponse.ReturnValue);
                this.subFolderIds.Add(createFolderResponse.FolderId);
                uint createfolderHandle = this.responseSOHTable[TestSuiteBase.FIRST][openFolderResponse.OutputHandleIndex];

                // Release Handle.
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopRelease, createfolderHandle, 0);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();
                payloadCount = 0;

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref this.pcxh,
                    PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.rgbOut,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable,
                    out payloadCount,
                    ref this.rgbAuxOut);
                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            }
            #endregion

            #region GetHierarchyTable
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopGetHierarchyTable, this.objHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopGetHierarchyTableResponse getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;
            Site.Assert.AreEqual<uint>(0, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTable should succeed and '0' is expected to be returned. The returned value is {0}.", getHierarchyTableResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][getHierarchyTableResponse.OutputHandleIndex];
            #endregion

            #region SetColumns
            // This call will verify requirements referring to "NoCompression" or "NoXorMagic" is set in the pulFlags.
            // Parameter auxInfo is no use for RopSetColumns command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSetColumns, this.objHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, setColumnsResponse.ReturnValue, "RopSetColumns should succeed and '0' is expected to be returned. The returned value is {0}.", setColumnsResponse.ReturnValue);
            #endregion

            #region QueryRows
            // Since a mass of folders are created in the interpersonal messages sub-tree folder on the server, the length of total folder name is so long that the server will return multiple extended buffers 
            // in rgbOut of EcDoRpcExt2 for RopQueryRows.
            // Moreover, because the payload is so big that server will compress it if the client doesn't request no compression through the flags NoCompression.
            // The server will also produce additional response data (i.e., there will be multiple extended buffers contained).
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopQueryRows, this.objHandle, ConstValues.MaximumRowCount);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            this.rgbOut = null;
            payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.Chain,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;
            Site.Assert.AreEqual<uint>(0, queryRowsResponse.ReturnValue, "RopQueryRows should succeed and '0' is expected to be returned. The returned value is {0}.", queryRowsResponse.ReturnValue);

            #region Verify the requirements related with Chain flag and packing
            this.ServerAddAdditionalData(payloadCount);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R642, the count of payload that ROP response contains is {0}.", payloadCount);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R642
            bool isVerifyR642 = payloadCount > 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR642,
                642,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pulFlags] If pulFlags contains Chain (0x00000004), the client allows chaining of ROP response payloads.");

            // The server will obfuscate the payload in method GetHierarchyTable and compress the payload in the method QueryRows.
            // If the code can reach here, this requirement will be verified.
            Site.CaptureRequirement(
                966,
                @"[In Extended Buffer Packing] The server can then compress and/or obfuscate this payload if the client requests and set the Flags field of the RPC_HEADER_EXT structure to indicate how the payload has been altered.");
            #endregion
            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case mainly verifies that ReadStream ROP commands can return multiple extended buffers.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC08_TestROPReadStream()
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Logon to mailbox
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
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region RopCreateMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateMessage, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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

            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, createMessageResponse.ReturnValue, "RopCreateMessage should succeed, the ReturnValue of its response is expected to be 0(success). The returned value is {0}.", createMessageResponse.ReturnValue);
            uint messageHandle = this.responseSOHTable[TestSuiteBase.FIRST][createMessageResponse.OutputHandleIndex];
            #endregion

            #region RopOpenStream
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, messageHandle, TestSuiteBase.OpenModeFlags);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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

            RopOpenStreamResponse openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, openStreamResponse.ReturnValue, "RopOpenStream should succeed, the ReturnValue of its response is expected to be 0(success). The returned value is {0}.", openStreamResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openStreamResponse.OutputHandleIndex];
            #endregion

            #region RopWriteStream
            RopWriteStreamResponse writeStreamResponse;
            for (int i = 0; i < int.Parse(Common.GetConfigurationPropertyValue("WriteStreamCount", this.Site)); i++)
            {
                // Parameter auxInfo is no use for RopWriteStream command, so set it to unUsedInfo
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopWriteStream, this.objHandle, this.unusedInfo);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();

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

                writeStreamResponse = (RopWriteStreamResponse)this.response;
                Site.Assert.AreEqual<uint>(0x00000000, writeStreamResponse.ReturnValue, "RopWriteStream should succeed, the ReturnValue of its response is expected to be 0(success). The returned value is {0}.", writeStreamResponse.ReturnValue);
            }
            #endregion

            #region RopCommitStream
            // Parameter auxInfo is no use for RopCommitStream command, so set it to unUsedInfo
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCommitStream, this.objHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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

            RopCommitStreamResponse commitStreamResponse = (RopCommitStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, commitStreamResponse.ReturnValue, "RopCommitStream should succeed, the ReturnValue of its response is expected to be 0(success). The returned value is {0}.", commitStreamResponse.ReturnValue);
            #endregion

            #region RopOpenStream
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, messageHandle, (ulong)ZERO); // Open the new committed stream for read-only access
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed for RopOpenStream and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openStreamResponse.ReturnValue, "RopOpenStream should succeed for RopOpenStream and '0' is expected to be returned. The returned value is {0}.", openStreamResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openStreamResponse.OutputHandleIndex];
            #endregion

            #region RopReadStream with requested small data
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopReadStream, this.objHandle, ConstValues.RequestedByteCount);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic | PulFlags.Chain,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 for small requested data should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopReadStreamResponse readStreamResponse = (RopReadStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0, readStreamResponse.ReturnValue, "RopReadStream for small requested data should succeed and '0' is expected to be returned. The returned value is {0}.", readStreamResponse.ReturnValue);

            // Add the debug information 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R972, the data size that server returned is {0}, the data size that client requested is {1}.", readStreamResponse.DataSize, ConstValues.RequestedByteCount);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R972
            Site.CaptureRequirementIfIsTrue(
                readStreamResponse.DataSize <= ConstValues.RequestedByteCount,
                972,
                @"[In Extended Buffer Packing]The server MUST NOT return more data to the client than the client originally requested.");

            #endregion

            #region RopReadStream with requested large enough data
            // An unsigned long indicates a value of maximumByteCount value in RopReadStream request, as specified by RopReadStream ROP in [MS-OXCROPS].
            ulong maximumByteWithLargeData = 327680;
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopReadStream, this.objHandle, maximumByteWithLargeData);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            uint payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic | PulFlags.Chain,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 for small requested data should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            readStreamResponse = (RopReadStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0, readStreamResponse.ReturnValue, "RopReadStream for small requested data should succeed and '0' is expected to be returned. The returned value is {0}.", readStreamResponse.ReturnValue);

            this.ServerAddAdditionalData(payloadCount);

            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case verifies that FastTransferSourceGetBuffer ROP commands can return multiple extended buffers.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC09_TestFastTransferSourceGetBuffer()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbOut = ConstValues.ValidpcbOut;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Logon to mailbox
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
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region RopCreateMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateMessage, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "RopCreateMessage should succeed and '0' is expected to be returned. The returned value is {0}.", createMessageResponse.ReturnValue);
            uint objCreateMessageHandle = this.responseSOHTable[TestSuiteBase.FIRST][createMessageResponse.OutputHandleIndex];
            #endregion

            #region RopOpenStream
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, objCreateMessageHandle, TestSuiteBase.OpenModeFlags);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopOpenStreamResponse openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, openStreamResponse.ReturnValue, "RopOpenStream should succeed, the ReturnValue of its response is expected to be 0(success). The actual value is {0}.", openStreamResponse.ReturnValue);
            uint objOpenStreamHandle = this.responseSOHTable[TestSuiteBase.FIRST][openStreamResponse.OutputHandleIndex];
            #endregion

            #region RopWriteStream
            RopWriteStreamResponse writeStreamResponse;
            uint count = uint.Parse(Common.GetConfigurationPropertyValue("WriteStreamCount", this.Site));
            for (int i = 0; i < count; i++)
            {
                // Parameter auxInfo is no use for RopWriteStream command, so set it to unUsedInfo.
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopWriteStream, objOpenStreamHandle, this.unusedInfo);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();

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

                writeStreamResponse = (RopWriteStreamResponse)this.response;
                Site.Assert.AreEqual<uint>(0x00000000, writeStreamResponse.ReturnValue, "RopWriteStream should succeed, the ReturnValue of its response is expected to be 0(success). The returned value is {0}.", writeStreamResponse.ReturnValue);
            }
            #endregion

            #region RopCommitStream
            // Parameter auxInfo is no use for RopCommitStream command, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCommitStream, objOpenStreamHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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

            RopCommitStreamResponse commitStreamResponse = (RopCommitStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0x00000000, commitStreamResponse.ReturnValue, "RopCommitStream should succeed, the ReturnValue of its response is expected to be 0(success). The actual value is {0}.", commitStreamResponse.ReturnValue);

            #endregion

            #region RopSaveChangesMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSaveChangesMessage, objCreateMessageHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, saveChangesMessageResponse.ReturnValue, "RopSaveChangesMessage should succeed and '0' is expected to be returned. The returned value is {0}.", saveChangesMessageResponse.ReturnValue);
            #endregion

            #region RopOpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and '0' is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openFolderResponse.OutputHandleIndex];
            #endregion

            #region RopFastTransferSourceCopyMessagesResponse
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopFastTransferSourceCopyMessages, this.objHandle, saveChangesMessageResponse.MessageId);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

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
            RopFastTransferSourceCopyMessagesResponse fastTransferSourceCopyMessagesResponse = (RopFastTransferSourceCopyMessagesResponse)this.response;
            Site.Assert.AreEqual<uint>(0, fastTransferSourceCopyMessagesResponse.ReturnValue, "RopFastTransferSourceCopyMessagesResponse should succeed and '0' is expected to be returned. The returned value is {0}.", fastTransferSourceCopyMessagesResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][fastTransferSourceCopyMessagesResponse.OutputHandleIndex];
            #endregion

            #region RopFastTransferSourceGetBuffer
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopFastTransferSourceGetBuffer, this.objHandle, ConstValues.MaximumBufferSize);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            uint payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic | PulFlags.Chain,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopFastTransferSourceGetBufferResponse fastTransferSourceGetBufferResponse = (RopFastTransferSourceGetBufferResponse)this.response;
            Site.Assert.AreEqual<uint>(0, fastTransferSourceGetBufferResponse.ReturnValue, "RopFastTransferSourceGetBuffer should succeed and '0' is expected to be returned. The returned value is {0}.", fastTransferSourceGetBufferResponse.ReturnValue);
            #endregion

            this.ServerAddAdditionalData(payloadCount);

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case verifies the results of each ROP command.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC10_TestMultipleROPs()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbOut = ConstValues.ValidpcbOut;
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Logon to Mailbox
            // This call will verify requirements referring to "NoCompression" or "NoXorMagic" set in the pulFlags.
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
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region RopOpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, this.objHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and '0' is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openFolderResponse.OutputHandleIndex];
            #endregion

            #region RopGetHierarchyTable
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopGetHierarchyTable, this.objHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            RopGetHierarchyTableResponse getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;
            Site.Assert.AreEqual<uint>(0, getHierarchyTableResponse.ReturnValue, "RopGetHierarchyTalbe should succeed and '0' is expected to be returned. The returned value is {0}.", getHierarchyTableResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][getHierarchyTableResponse.OutputHandleIndex];
            #endregion

            #region Call multiple ROP commands on one EcDoRpcExt2 method
            // Parameter auxInfo is no use for composing multiple ROP commands once, so set it to unUsedInfo.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.MultipleRops, this.objHandle, this.unusedInfo);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            this.rgbOut = null;
            uint payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed for multiple ROP commands and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            List<IDeserializable> ropResponse = this.ParseMultipleRopsResponse(this.rgbOut);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R621, the response data for RopSetColumns is {0}, the response data for RopQueryRows is {1}.", ((RopSetColumnsResponse)ropResponse[0]).RopId, ((RopQueryRowsResponse)ropResponse[1]).RopId);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R621
            // ropResponse[0] is the response data for RopSetColumns and ropResponse[1] is the response data for RopQueryRows
            bool isVerifyR621 =
                (ropResponse.Count > 1) &&
                (((RopSetColumnsResponse)ropResponse[0]).RopId != ((RopQueryRowsResponse)ropResponse[1]).RopId);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR621,
                621,
                @"[In EcDoRpcExt2 Method (opnum 11)] The server returns the results of each ROP command to the client.");

            #endregion Capture R621

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This test case mainly tests the requirements related to the Reserved field in the AUX_PERF_SESSIONINFO, AUX_PERF_SESSIONINFO_V2, 
        /// AUX_PERF_CLIENTINFO, AUX_PERF_DEFMDB_SUCCESS, AUX_PERF_DEFGC_SUCCESS, AUX_PERF_MDB_SUCCESS_V2, AUX_PERF_GC_SUCCESS_V2, 
        /// AUX_PERF_FAILURE, AUX_PERF_ACCOUNTINFO and AUX_CLIENT_CONNECTION_INFO structures. 
        /// It also tests the requirements related to Reserved_1 and Reserved_2 fields in the AUX_PERF_PROCESSINFO, 
        /// AUX_PERF_GC_SUCCESS and AUX_PERF_FAILURE_V2 structures. 
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC11_TestReserved()
        {
            this.CheckTransport();

            uint resultOneOfEcDoConnectEx;
            uint resultTwoOfEcDoConnectEx;
            uint resultOneOfEcDoRpcExt2;
            uint resultTwoOfEcDoRpcExt2;
            byte[] payload;

            // An integer indicates a value for reserved field, as specified by EcDoConnectEx method and EcDoRpcExt2 method in [MS-OXCRPC].
            int reserveValue1 = 1;

            // An integer indicates a value for reserved field, as specified by EcDoConnectEx method and EcDoRpcExt2 method in [MS-OXCRPC].
            int reserveValue2 = 2;

            #region Send AUX_PERF_SESSIONINFO structure to server.
            payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(reserveValue1);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_SESSIONINFO, payload);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_SESSIONINFO, payload);

            payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(reserveValue2);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_SESSIONINFO, payload);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_SESSIONINFO, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1819");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1819
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1819,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_SESSIONINFO.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1825");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1825
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1825,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_SESSIONINFO.");
            #endregion

            #region Send AUX_PERF_SESSIONINFO_V2 structure to server
            payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO_V2(reserveValue1);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);

            payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO_V2(reserveValue2);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1820");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1820
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1820,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_SESSIONINFO_V2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1826");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1826
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1826,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_SESSIONINFO_V2.");
            #endregion

            #region Send AUX_PERF_CLIENTINFO to server
            // The ClientMode 0x00 means that client is not designating a mode of operation.
            payload = AdapterHelper.Compose_AUX_PERF_CLIENTINFO(reserveValue1, 0x00);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_CLIENTINFO, payload);

            // The ClientMode 0x00 means that client is not designating a mode of operation.
            payload = AdapterHelper.Compose_AUX_PERF_CLIENTINFO(reserveValue2, 0x00);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_CLIENTINFO, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1821");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1821
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1821,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_CLIENTINFO.");

            // The ClientMode 0x01 means that client is running in classic online mode.
            payload = AdapterHelper.Compose_AUX_PERF_CLIENTINFO(reserveValue1, 0x01);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_CLIENTINFO, payload);

            // The ClientMode 0x02 means that client is running in cached mode.
            payload = AdapterHelper.Compose_AUX_PERF_CLIENTINFO(reserveValue1, 0x02);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_CLIENTINFO, payload);
            #endregion

            #region Send AUX_PERF_PROCESSINFO structure to server
            payload = AdapterHelper.Compose_AUX_PERF_PROCESSINFO(reserveValue1, ReserveDefault);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_PROCESSINFO, payload);
            payload = AdapterHelper.Compose_AUX_PERF_PROCESSINFO(reserveValue2, ReserveDefault);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_PROCESSINFO, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1822");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1822
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1822,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_1 field in structure AUX_PERF_PROCESSINFO.");

            payload = AdapterHelper.Compose_AUX_PERF_PROCESSINFO(ReserveDefault, reserveValue1);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_PROCESSINFO, payload);
            payload = AdapterHelper.Compose_AUX_PERF_PROCESSINFO(ReserveDefault, reserveValue2);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_PERF_PROCESSINFO, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1823");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1823
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1823,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_2 field in structure AUX_PERF_PROCESSINFO.");
            #endregion

            #region Send AUX_PERF_DEFMDB_SUCCESS structure to server
            payload = AdapterHelper.Compose_AUX_PERF_DEFMDB_SUCCESS(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_DEFMDB_SUCCESS, payload);
            payload = AdapterHelper.Compose_AUX_PERF_DEFMDB_SUCCESS(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_DEFMDB_SUCCESS, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1827");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1827
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1827,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_DEFMDB_SUCCESS.");
            #endregion

            #region Send AUX_PERF_DEFGC_SUCCESS structure to server
            payload = AdapterHelper.Compose_AUX_PERF_DEFGC_SUCCESS(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_DEFGC_SUCCESS, payload);
            payload = AdapterHelper.Compose_AUX_PERF_DEFGC_SUCCESS(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_DEFGC_SUCCESS, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1828");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1828
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1828,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_DEFGC_SUCCESS.");
            #endregion

            #region Send AUX_PERF_MDB_SUCCESS_V2 structure to server
            payload = AdapterHelper.Compose_AUX_PERF_MDB_SUCCESS_V2(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_MDB_SUCCESS_V2, payload);
            payload = AdapterHelper.Compose_AUX_PERF_MDB_SUCCESS_V2(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_MDB_SUCCESS_V2, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1829");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1829
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1829,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_MDB_SUCCESS_V2.");
            #endregion

            #region Send AUX_PERF_GC_SUCCESS structure to server
            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS(reserveValue1, ReserveDefault);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS, payload);
            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS(reserveValue2, ReserveDefault);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1830");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1830
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1830,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_1 field in structure AUX_PERF_GC_SUCCESS.");

            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS(ReserveDefault, reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS, payload);
            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS(ReserveDefault, reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1831");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1831
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1831,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_2 field in structure AUX_PERF_GC_SUCCESS.");
            #endregion

            #region Send AUX_PERF_GC_SUCCESS_V2 structure to server
            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS_V2(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS_V2, payload);
            payload = AdapterHelper.Compose_AUX_PERF_GC_SUCCESS_V2(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_GC_SUCCESS_V2, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1832");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1832
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1832,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_GC_SUCCESS_V2.");
            #endregion

            #region Send AUX_PERF_FAILURE structure to server
            payload = AdapterHelper.Compose_AUX_PERF_FAILURE(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE, payload);
            payload = AdapterHelper.Compose_AUX_PERF_FAILURE(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1833");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1833
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1833,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_PERF_FAILURE.");
            #endregion

            #region Send AUX_PERF_FAILURE_V2 structure to server
            payload = AdapterHelper.Compose_AUX_PERF_FAILURE_V2(reserveValue1, ReserveDefault);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE_V2, payload);
            payload = AdapterHelper.Compose_AUX_PERF_FAILURE_V2(reserveValue2, ReserveDefault);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE_V2, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1834");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1834
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1834,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_1 field in structure AUX_PERF_FAILURE_V2.");

            payload = AdapterHelper.Compose_AUX_PERF_FAILURE_V2(ReserveDefault, reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE_V2, payload);
            payload = AdapterHelper.Compose_AUX_PERF_FAILURE_V2(ReserveDefault, reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_FAILURE_V2, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1835");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1835
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoRpcExt2,
                resultTwoOfEcDoRpcExt2,
                1835,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the Reserved_2 field in structure AUX_PERF_FAILURE_V2.");
            #endregion

            #region Send AUX_PERF_ACCOUNTINFO structure to server
            payload = AdapterHelper.Compose_AUX_PERF_ACCOUNTINFO(reserveValue1);
            resultOneOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_ACCOUNTINFO, payload);
            payload = AdapterHelper.Compose_AUX_PERF_ACCOUNTINFO(reserveValue2);
            resultTwoOfEcDoRpcExt2 = this.SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum.AUX_PERF_ACCOUNTINFO, payload);
            Site.Assert.AreEqual<uint>(resultOneOfEcDoRpcExt2, resultTwoOfEcDoRpcExt2, "The Reserved field in AUX_PERF_ACCOUNTINFO should be ignored.");
            #endregion

            #region Send AUX_CLIENT_CONNECTION_INFO structure to server
            // The ConnectionFlags value of 0x0001 for this field means that the client is running in cached mode.
            payload = AdapterHelper.Compose_AUX_CLIENT_CONNECTION_INFO(reserveValue1, 0x0001);
            resultOneOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_CLIENT_CONNECTION_INFO, payload);

            // The ConnectionFlags value of 0x0001 for this field means that the client is running in cached mode.
            payload = AdapterHelper.Compose_AUX_CLIENT_CONNECTION_INFO(reserveValue2, 0x0001);
            resultTwoOfEcDoConnectEx = this.SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum.AUX_CLIENT_CONNECTION_INFO, payload);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1824");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1824
            Site.CaptureRequirementIfAreEqual<uint>(
                resultOneOfEcDoConnectEx,
                resultTwoOfEcDoConnectEx,
                1824,
                @"[In Processing Auxiliary Buffers Received from the Client] Reply is the same for two different values used for the reserved field in structure AUX_CLIENT_CONNECTION_INFO.");
            #endregion
        }

        /// <summary>
        /// This case tests the server supports all functionality from previous server version level 
        /// when the functionality is supported at one server version level.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC12_TestServerVersionFunctionality()
        {
            this.CheckTransport();

            #region Client connects with Server
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            ushort[] rgwServerVersion;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out rgwServerVersion,
                out this.rgwBestVersion,
                out this.picxr,
                1);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx for server version test should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCRPC_R1078, the rgwServerVersion is {0}",
                rgwServerVersion[0].ToString() + "," + rgwServerVersion[1].ToString() + "," + rgwServerVersion[2].ToString());

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1078
            // The rgwServerVersion contains three words.
            // If server returns the server version then at least one word of rgwServerVersion must not be zero.
            bool isVerifyR1078 = rgwServerVersion[0] != 0 || rgwServerVersion[1] != 0 || rgwServerVersion[2] != 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1078,
                1078,
                @"[In Version Checking] When the server receives the client version in the EcDoConnectEx method, the server returns its [server's] version to the client.");

            bool isSupported = this.IsFunctionalitySupported(rgwServerVersion, this.pcxh);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1115, server does {0} support all functionality from previous server version levels.", isSupported ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1115
            // According to the stack implementation of method IsServerVersionSupportCheck, if this method returns true, 
            // it means server supports the corresponding features.
            Site.CaptureRequirementIfIsTrue(
                isSupported,
                1115,
                @"[In Server Versions] To support functionality at a given server version level, the server MUST support all functionality from previous server version levels.");
            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case tests whether the specific authentication method is supported.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC13_TestAuthenticationMethods()
        {
            this.CheckTransport();

            #region Tests whether the RPC_C_AUTHN_WINNT is supported.
            if (Common.IsRequirementEnabled(1749, this.Site))
            {
                this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, (uint)AuthenticationService.RPC_C_AUTHN_WINNT, this.userName, this.password);
                Site.Assert.IsTrue(this.returnStatus, "Initializing the RPC binding should succeed.");

                this.pcbOut = ConstValues.ValidpcbOut;
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1749");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1749
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1749,
                    @"[In Appendix B: Product Behavior] Implementation does support RPC_C_AUTHN_WINNT authentication method. (Microsoft Exchange Server 2007 and above follow this behavior.)");

                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Assert.AreEqual<uint>(0, this.returnValue, "EcDodisconnect should succeed.");
            }
            #endregion

            #region Tests whether the RPC_C_AUTHN_GSS_KERBEROS is supported.
            if (Common.IsRequirementEnabled(1750, this.Site))
            {
                this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, (uint)AuthenticationService.RPC_C_AUTHN_GSS_KERBEROS, this.userName, this.password);
                Site.Assert.IsTrue(this.returnStatus, "Initializing the RPC binding should succeed.");

                this.pcbOut = ConstValues.ValidpcbOut;
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1750");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1750
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1750,
                    @"[In Appendix B: Product Behavior] Implementation does support RPC_C_AUTHN_GSS_KERBEROS authentication methods. (Microsoft Exchange Server 2007 and above follow this behavior.)");

                if (Common.IsRequirementEnabled(1915, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1915");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1915
                    // Test case use "exchangeMDB/<Mailbox server FQDN>" as SPN for the Kerberos authentication method.
                    // If all of the above steps succeed, R1915 will be verified.
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0,
                        this.returnValue,
                        1915,
                        @"[In Appendix B: Product Behavior] Implementation does use ""exchangeMDB/<Mailbox server FQDN>"" as the service principal name (SPN) for the Kerberos authentication method. (Exchange 2007 and above follow this behavior.)");
                }

                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Assert.AreEqual<uint>(0, this.returnValue, "EcDodisconnect should succeed.");
            }
            #endregion

            #region Tests whether the RPC_C_AUTHN_GSS_NEGOTIATE is supported.
            if (Common.IsRequirementEnabled(1751, this.Site))
            {
                this.returnStatus = this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, (uint)AuthenticationService.RPC_C_AUTHN_GSS_NEGOTIATE, this.userName, this.password);
                Site.Assert.IsTrue(this.returnStatus, "Initializing the RPC binding should succeed.");

                this.pcbOut = ConstValues.ValidpcbOut;
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1751");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1751
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1751,
                    @"[In Appendix B: Product Behavior] Implementation does support RPC_C_AUTHN_GSS_NEGOTIATE authentication method. (Microsoft Exchange Server 2007 and above follow this behavior.)");

                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Assert.AreEqual<uint>(0, this.returnValue, "EcDodisconnect should succeed.");
            }
            #endregion

            #region Tests whether the RPC_C_AUTHN_NONE is supported.
            if (Common.IsRequirementEnabled(1550, this.Site))
            {
                this.returnStatus = this.oxcrpcAdapter.InitializeRPC((uint)AuthenticationLevel.RPC_C_AUTHN_LEVEL_NONE, (uint)AuthenticationService.RPC_C_AUTHN_NONE, this.userName, this.password);
                Site.Assert.IsTrue(this.returnStatus, "Initializing the RPC binding should succeed.");

                this.pcbOut = ConstValues.ValidpcbOut;
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

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1550");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1550
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1550,
                    @"[In Appendix B: Product Behavior] Implementation does support the RPC_C_AUTHN_NONE authentication method. (Microsoft Exchange Server 2013 and above follow this behavior.)");

                this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
                Assert.AreEqual<uint>(0, this.returnValue, "EcDodisconnect should succeed.");
            }
            #endregion
        }

        /// <summary>
        /// This case tests the error code of EcDoConnect method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC14_TestErrorCodeOfEcDoConnect()
        {
            this.CheckTransport();

            #region Tests error code ecAccessDenied
            if (Common.IsRequirementEnabled(4887, this.Site))
            {
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                    ref this.pcxh,
                    TestSuiteBase.UlIcxrLinkForNoSessionLink,
                    ref this.pulTimeStamp,
                    null,
                    string.Empty,
                    ref this.pcbAuxOut,
                    this.rgwClientVersion,
                    out this.rgwBestVersion,
                    out this.picxr);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4887");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R4887
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    this.returnValue,
                    4887,
                    @"[In Appendix B: Product Behavior] Implementation does return ecAccessDenied (0x80070005) if the szUserDN parameter is empty. (<15> Section 3.1.4.1: Exchange 2010 and above follow this behavior.)");
            }

            ushort[] rgwServerVersion;

            if (Common.IsRequirementEnabled(1437, this.Site))
            {
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                uint flagsValue = 0x00000000;

                this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                    ref this.pcxh,
                    TestSuiteBase.UlIcxrLinkForNoSessionLink,
                    ref this.pulTimeStamp,
                    null,
                    string.Empty,
                    ref this.pcbAuxOut,
                    this.rgwClientVersion,
                    out rgwServerVersion,
                    out this.rgwBestVersion,
                    out this.picxr,
                    flagsValue);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1437");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1437
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000000,
                    this.returnValue,
                    1437,
                    @"[In Appendix B: Product Behavior] [<15> Section 3.1.4.1] Implementation returns ecNone (0x00000000) if the szUserDN parameter is empty. (Microsoft Exchange Server 2007 follows this behavior.)");
            }
            #endregion

            #region Tests error code ecUnknownUser
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                "UserDN",
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R600");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R600
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000003EB,
                this.returnValue,
                600,
                @"[In EcDoConnectEx Method (Opnum 10)] [Return Values] [ecUnknownUser (0x000003EB)] The server does not recognize the szUserDN parameter as a valid enabled mailbox.");
            #endregion

            #region Tests error code ecVersionMismatch
            ushort[] mismatchClientVersion = new ushort[] { 0, 0, 0 };

            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
             ref this.pcxh,
             TestSuiteBase.UlIcxrLinkForNoSessionLink,
             ref this.pulTimeStamp,
             null,
             this.userDN,
             ref this.pcbAuxOut,
             mismatchClientVersion,
             out this.rgwBestVersion,
             out this.picxr);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R606");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R606
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040110,
                this.returnValue,
                606,
                @"[In EcDoConnectEx Method (Opnum 10)] [Return Values] [ecVersionMismatch (0x80040110)] The client and server versions are not compatible.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R608");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R608
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040110,
                this.returnValue,
                608,
                @"[In EcDoConnectEx Method (Opnum 10)] [Return Values] [ecVersionMismatch (0x80040110)] The client protocol version is earlier than that required by the server.");

            #endregion

            #region Tests error code ecClientVerDisallowed.
            List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                mismatchClientVersion,
                out rgwServerVersion,
                out this.rgwBestVersion,
                out this.picxr,
                0x00008000,
                out rgbAuxOutValue);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4886");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R4886
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0x000004DF,
                this.returnValue,
                4886,
                @"[In Client Versions] [Client version ""12.00.0000.000""] For client versions earlier than 12.00.0000.000, the server MUST not fail the EcDoConnectEx method call with ecClientVerDisallowed if the EcDoConnectEx method parameter flag 0x00008000 is passed in the ulFlags parameter.");
            #endregion
        }

        /// <summary>
        /// This case tests the AUX_OSVERSIONINFO structure.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC15_TestAUXOSVERSIONINFO()
        {
            this.CheckTransport();

            List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue;
            ushort[] rgwServerVersion;

            #region Client connects with Server
            byte[] payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(ReserveDefault);
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(RgbAuxInEnum.AUX_PERF_SESSIONINFO, payload);
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out rgwServerVersion,
                out this.rgwBestVersion,
                out this.picxr,
                0x00000001,
                out rgbAuxOutValue);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoConnectEx should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Verify the AUX_OSVERSIONINFO structure.
            if (Common.IsRequirementEnabled(1918, this.Site))
            {
                foreach (AUX_SERVER_TOPOLOGY_STRUCTURE rgbAuxValue in rgbAuxOutValue)
                {
                    if (rgbAuxValue.Header.Version == 0x01 && rgbAuxValue.Header.Type == 0x16)
                    {
                        string[] operatingSystemVersions = this.oxcrpcControlAdapter.GetOSVersions().Split(new string[] { "." }, StringSplitOptions.None);
                        Site.Assert.IsTrue(operatingSystemVersions.Length == 5, "Operating system version format should be valid.");

                        int index = 0;
                        int operatingSystemVersionInfoSize = BitConverter.ToInt32(rgbAuxValue.Payload, 0);
                        index += 4;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R272");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R272
                        Site.CaptureRequirementIfAreEqual<int>(
                            rgbAuxValue.Payload.Length,
                            operatingSystemVersionInfoSize,
                            272,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] OSVersionInfoSize (4 bytes): The size of this AUX_OSVERSIONINFO structure.");

                        int majorVersion = BitConverter.ToInt32(rgbAuxValue.Payload, index);
                        index += 4;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R273");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R273
                        Site.CaptureRequirementIfAreEqual<int>(
                            int.Parse(operatingSystemVersions[0]),
                            majorVersion,
                            273,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] MajorVersion (4 bytes): The major version number of the operating system of the server.");

                        int minorVersion = BitConverter.ToInt32(rgbAuxValue.Payload, index);
                        index += 4;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R274");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R274
                        Site.CaptureRequirementIfAreEqual<int>(
                            int.Parse(operatingSystemVersions[1]),
                            minorVersion,
                            274,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] MinorVersion (4 bytes): The minor version number of the operating system of the server.");

                        int buildNumber = BitConverter.ToInt32(rgbAuxValue.Payload, index);
                        index += 4;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R275");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R275
                        Site.CaptureRequirementIfAreEqual<int>(
                            int.Parse(operatingSystemVersions[2]),
                            buildNumber,
                            275,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] BuildNumber (4 bytes): The build number of the operating system of the server.");

                        index += 132;

                        short servicePackMajor = BitConverter.ToInt16(rgbAuxValue.Payload, index);
                        index += 2;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R278");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R278
                        Site.CaptureRequirementIfAreEqual<short>(
                            short.Parse(operatingSystemVersions[3]),
                            servicePackMajor,
                            278,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] ServicePackMajor (2 bytes): The major version number of the latest operating system service pack that is installed on the server.");

                        short servicePackMinor = BitConverter.ToInt16(rgbAuxValue.Payload, index);

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R279");

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R279
                        Site.CaptureRequirementIfAreEqual<short>(
                            short.Parse(operatingSystemVersions[4]),
                            servicePackMinor,
                            279,
                            @"[In AUX_OSVERSIONINFO Auxiliary Block Structure] ServicePackMinor (2 bytes): The minor version number of the latest operating system service pack that is installed on the server.");

                        break;
                    }
                }
            }
            #endregion

            #region Call EcDoDisconnect to destroy the Session Context on the server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case verifies the requirements related to EcRRegisterPushNotification method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC16_TestEcRRegisterPushNotification()
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoConnectEx method should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcRRegisterPushNotification
            string callBackAddress = Common.GetConfigurationPropertyValue("NotificationIP", this.Site);

            // An integer indicates a valid value of cbContext that should be less than or equal to 16 (0x10), as specified by EcRRegisterPushNotification method in [MS-OXCRPC].
            int validcbContext = 16;

            uint outHinder;
            bool isVerifyR379 = false;
            byte[] rgbContext = new byte[validcbContext];

            if (Common.IsRequirementEnabled(1558, this.Site))
            {
                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.AF_INET, callBackAddress, out outHinder);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1558");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1558
                // According to the Error Codes of EcRRegisterPushNotification specified in Open Specification,
                // the value of the ecNotSupported is 0x80040102.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    this.returnValue,
                    1558,
                    @"[In Appendix B: Product Behavior] Implementation does not support the EcRRegisterPushNotification method call. <27> Section 3.1.4.5: Exchange 2013 and Exchange 2016 do not support the EcRRegisterPushNotification RPC and always returns ecNotSupported.");
            }

            if (Common.IsRequirementEnabled(1845, this.Site))
            {
                #region Verify requirements related with error code ecTooBig
                // An integer indicates an invalid value of cbContext that should be larger than 16 (0x10), as specified by EcRRegisterPushNotification method in [MS-OXCRPC].
                int tooBigcbContext = 20;

                // Call EcRRegisterPushNotification method with too big opaque context data.
                rgbContext = new byte[tooBigcbContext];
                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.AF_INET, callBackAddress, out outHinder);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R438");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R438
                // Because the length of rgbContext is 0x00000014 that is larger than 0x00000010.
                // If server fails EcRRegisterPushNotification method call with error code 0x80040305, then R438 will be verified.
                Site.CaptureRequirementIfAreEqual<uint>(
                   0x80040305,
                   this.returnValue,
                   438,
                   @"[In EcRRegisterPushNotification Method (opnum 4)] [cbContext] If the value of this parameter is larger than 0x00000010, the server MUST fail the call with error code ecTooBig.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R459");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R459
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040305,
                    this.returnValue,
                    459,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [Return Values] ecTooBig (0x80040305) means ""Opaque context data is too large"".");
                #endregion

                #region Verify requirements related with error code ecInvalidParam
                rgbContext = new byte[validcbContext];

                // Call EcRRegisterPushNotification method with cbCallbackAddress that
                // does not correspond to the sockaddr size based on address family.
                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.SIZENOTCORRESPONDSOCKADDRSIZE, callBackAddress, out outHinder);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R450");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R450
                // The Add_Families is set to not correspond to the sockaddr size based on address family, 
                // so the size will be Too Big. The value of the ecInvalidParam 0x80070057.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.returnValue,
                    450,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [cbCallbackAddress] If this size does not correspond to the size of the sockaddr structure based on address family, the server MUST return error code ecInvalidParam.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R455");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R455
                // Call EcRRegisterPushNotification method with invalid cbCallbackAddress parameter.
                // So if server fails with 0x80070057 then R455 will be verified.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.returnValue,
                    455,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [Return Values] ecInvalidParam (0x80070057) means ""A parameter passed was not valid for the call"".");

                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.NOT_SUPPORTED, callBackAddress, out outHinder);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R446");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R446
                // The address family has been set to "NOT_SUPPORTED" already, here just verify the error code. The value of ecInvalidParam is 0x80070057.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.returnValue,
                    446,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [rgbCallbackAddress] If an address family is requested that is not supported, the server MUST return error code ecInvalidParam.");

                #endregion Verify requirements related with error code ecInvalidParam

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
                Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);
                this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
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

                #region Call EcRRegisterPushNotification when pcxh is valid
                string rgbContextValue = "Notification";
                rgbContext = ASCIIEncoding.ASCII.GetBytes(rgbContextValue);

                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.AF_INET, callBackAddress, out outHinder);

                #endregion

                #region Capture code.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R444");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R444
                // Since the EcRRegisterPushNotification uses the Add_Families.AF_INET, here only checks the return value. If it returns success, it means the AF_INET is supported.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    444,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [rgbCallbackAddress] The server supports the address families AF_INET for a callback address that corresponds to the protocol sequence types that are specified in section 2.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R452");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R452
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    452,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] Return Values: If the method succeeds, the return value is 0.");

                #endregion 

                #region Receive push notification on the specified port.
                // Trigger the event
                bool isCreateMailSuccess = this.oxcrpcControlAdapter.CreateMailItem();
                Site.Assert.IsTrue(isCreateMailSuccess, "CreateMailItem method should execute successfully.");

                string opaqueReturned = this.GetPushNotification(Add_Families.AF_INET);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R415, the method EcRRegisterPushNotification does {0} register a callback address with the server for a Session Context.", string.IsNullOrEmpty(opaqueReturned) ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R415
                // Because the EcRRegisterPushNotification has registered the callback address with server.
                // And if opaqueReturned is not null, then server successfully notifies the client of pending events.
                // So R415 will be verified.
                Site.CaptureRequirementIfIsNotNull(
                    opaqueReturned,
                    415,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] The EcRRegisterPushNotification method registers a callback address with the server for a Session Context.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R436");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R436
                Site.CaptureRequirementIfAreEqual<string>(
                    rgbContextValue,
                    opaqueReturned,
                    436,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [rgbContext] The server MUST save this data within the Session Context and use it when sending a notification to the client.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R422, the server does {0} send a packet containing just the opaque context data to the callback address.", string.IsNullOrEmpty(opaqueReturned) ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R422
                // Because the EcRRegisterPushNotification has registered the callback address with server.
                // And if opaqueReturned is not null, then server successfully notifies the client of pending events.
                // So R422 will be verified.
                Site.CaptureRequirementIfIsNotNull(
                    opaqueReturned,
                    422,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] To notify the client of pending events, the server sends a packet containing just the opaque context data to the callback address.");
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
                #endregion

                #region Call EcRRegisterPushNotification when pcxh is invalid
                this.pcxhInvalid = (IntPtr)ConstValues.InvalidPcxh;
                this.returnValueForInvalidCXH = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxhInvalid, rgbContext, Add_Families.AF_INET, callBackAddress, out outHinder);
                Site.Assert.AreNotEqual<uint>(0, this.returnValueForInvalidCXH, "EcRRegisterPushNotification should not succeed by using an invalid CXH and '0' is not expected to be returned. The returned value is {0}.", this.returnValueForInvalidCXH);
                #endregion

                #region Capture code
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R379, the return value of method EcRRegisterPushNotification with parameter pcxh valid is {0}, the return value of method EcRRegisterPushNotification with parameter pcxh invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R379
                // The EcRRegisterPushNotification method invoked above uses the active CXH returned form EcDoConnectEx, so if the code can reach here, the requirement is verified.
                isVerifyR379 = (this.returnValue == TestSuiteBase.ResultSuccess) && (this.returnValueForInvalidCXH != TestSuiteBase.ResultSuccess);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR379,
                    379,
                    @"[In Message Processing Events and Sequencing Rules] [EcRRegisterPushNotification] The method requires an active session context handle to be returned from the EcDoConnectEx method.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R417, the return value of method EcRRegisterPushNotification with parameter pcxh valid is {0}, the return value of method EcRRegisterPushNotification with parameter pcxh invalid is {1}.", this.returnValue, this.returnValueForInvalidCXH);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R417
                // Since the context R417 is similar with the requirement 379, if R379 is verified, it means R417 is verified.
                bool isVerifyR417 = isVerifyR379;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR417,
                    417,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] This [the method EcRRegisterPushNotification] call requires an active session context handle from the EcDoConnectEx method, as specified in section 3.1.4.1.");

                callBackAddress = Common.GetConfigurationPropertyValue("NotificationIPv6", this.Site);
                this.returnValue = this.oxcrpcAdapter.EcRRegisterPushNotification(ref this.pcxh, rgbContext, Add_Families.AF_INET6, callBackAddress, out outHinder);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1301");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1301
                // If this call returns success (return == 0), it means the server supports the address families AF_INET6
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.returnValue,
                    1301,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [rgbCallbackAddress] The server supports the address families AF_INET6 for a callback address that corresponds to the protocol sequence types that are specified in section 2.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1845, the return value of calling method EcRRegisterPushNotification is {0}.", this.returnValue);

                // If code can run here, it indicates that Exchange 2007 supports method EcRRegisterPushNotification.
                Site.CaptureRequirement(
                    1845,
                    @"[In Appendix B: Product Behavior] Implementation does support method EcRRegisterPushNotification. <7> Section 3.1.4: Exchange 2007 supports method EcRRegisterPushNotification.");

                #endregion
            }
            #endregion

            #region Call EcDoDisconnect with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case verifies the requirements related to EcDoAsyncConnectEx method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC17_TestEcDoAsyncConnectEx()
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoConnectEx method should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoAsyncConnectEx with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxh, ref this.pacxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoAsyncConnectEx should succeed by using a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoAsyncConnectEx when pcxh is invalid
            this.pcxhInvalid = (IntPtr)ConstValues.InvalidPcxh;
            IntPtr pacxhForInvalidCXH = new IntPtr();
            this.returnValueForInvalidCXH = this.oxcrpcAdapter.EcDoAsyncConnectEx(this.pcxhInvalid, ref pacxhForInvalidCXH);
            Site.Assert.AreNotEqual<uint>(0, this.returnValueForInvalidCXH, "EcDoAsyncConnectEx should not succeed by using an invalid CXH and '0' is not expected to be returned. The returned value is {0}.", this.returnValueForInvalidCXH);
            #endregion

            #region Capture code.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R705, the return value of method EcDoAsyncConnectEx with parameter CXH returned from EcDoConnectEx is {0}, the return value of method EcDoAsyncConnectEx with parameter CXH not returned from EcDoConnectEx is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R705
            // When the CXH is returned from EcDoConnectEx method, the call is successful. While the CXH is not returned from the EcDoConnectEx method, the call is failed. 
            // So if the code can reach here, the requirement is verified.
            bool isVerifyR705 = (this.returnValue == ResultSuccess) && (this.returnValueForInvalidCXH != 0);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR705,
                705,
                @"[In EcDoAsyncConnectEx Method (opnum 14)] This [method EcDoAsyncConnectEx] call requires that an active session context handle be returned from the EcDoConnectEx method.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R398, the return value of method EcDoAsyncConnectEx with parameter CXH returned from EcDoConnectEx is {0}, the return value of method EcDoAsyncConnectEx with parameter CXH not returned from EcDoConnectEx is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R398
            // Since the context of R398 is similar with the requirement R705, if R705 is verified, it means that this requirement will be verified
            bool isVerifyR398 = isVerifyR705;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR398,
                398,
                @"[In Message Processing Events and Sequencing Rules] [EcDoAsyncConnectEx] The method requires an active session context handle to be returned from the EcDoConnectEx method.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R711");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R711
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                pacxhForInvalidCXH.ToInt32(),
                711,
                @"[In EcDoAsyncConnectEx Method (opnum 14)] [pacxh] On failure the returned value is NULL.");
            #endregion

            #region Call EcDoDisconnect with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }

        /// <summary>
        /// This case verifies the requirements related to EcDoRpcExt2 method.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC18_TestEcDoRpcExt2()
        {
            this.CheckTransport();

            byte[] payload = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(ReserveDefault);
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2, payload);

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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoConnectEx method should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDummyRpc
            this.returnValue = this.oxcrpcAdapter.EcDummyRpc();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R464");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R464
            // Server returns 0 means that calling EcDummyRpc is successful.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                464,
                @"[In EcDummyRpc Method (opnum 6)] The EcDummyRpc method returns a SUCCESS.");
            #endregion

            #region Call EcDoRpcExt2 with pulFlags contains NoXorMagic.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.rgbAuxOut = new byte[this.pcbAuxOut];
            uint payloadCount = 0;
            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                this.rgbAuxIn,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            #region Capture code.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R694");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R694
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                694,
                @"[In EcDoRpcExt2 Method (opnum 11)] Return Values: If the method succeeds, the return value is 0.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R475");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R475
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.returnValue,
                475,
                @"[In EcDoConnectEx Method (Opnum 10)] The EcDoConnectEx method establishes a new Session Context with the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1290, the value of server return is {0}", this.returnValue);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1290
            // Calls EcDoRpcExt2 method and passes a RopLogon command to server. 
            // If server returns 0 and rgbOut contains the RopLogonResponse data then R1290 will be verified.
            bool isVerifyR1290 = this.returnValue == 0 && this.response is RopLogonResponse;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1290,
                1290,
                @"[In EcDoRpcExt2 Method (opnum 11)] The EcDoRpcExt2 method passes generic ROP commands to the server for processing within a Session Context.");
            #endregion

            #region Verify requirements related with NoXorMagic flag
            short flagsInRgbOut = BitConverter.ToInt16(this.rgbOut, 2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R638, the value of Flags in RPC_HEADER_EXT is {0}", flagsInRgbOut);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R638
            // Calling EcDoRpcExt2 with RopLogon ROP command only returns one RPC_HEADER_EXT.
            // R638 will be verified if flags in RPC_HEADER_EXT doesn't contain 0x0002, because NoXorMagic is contained in pulFlags when calling EcEoRpcExt2.
            bool isVerifyR638 = (flagsInRgbOut & (short)RpcHeaderExtFlags.XorMagic) != (short)RpcHeaderExtFlags.XorMagic && payloadCount == 1;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR638,
                638,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pulFlags] If pulFlags contains NoXorMagic (0x00000002), the server MUST NOT obfuscate the ROP response payload (rgbOut).");

            if (this.rgbAuxOut.Length != 0)
            {
                short flagsInRgbAuxOut = BitConverter.ToInt16(this.rgbAuxOut, 2);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1305, the value of Flags in RPC_HEADER_EXT is {0}", flagsInRgbAuxOut);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1305
                // R1305 will be verified if flags in RPC_HEADER_EXT doesn't contain 0x0002, because NoXorMagic is contained in pulFlags when calling EcDoRpcExt2.
                bool isVerifyR1305 = (flagsInRgbAuxOut & (short)RpcHeaderExtFlags.XorMagic) != (short)RpcHeaderExtFlags.XorMagic;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1305,
                    1305,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [pulFlags] If pulFlags contains NoXorMagic (0x00000002), the server MUST NOT obfuscate the auxiliary payload (rgbAuxOut).");
            }
            #endregion
            #endregion

            #region Call EcDoDisconnect with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
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
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx for no session context linking should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 with pulFlags contains NoCompression.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.rgbAuxOut = new byte[this.pcbAuxOut];
            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                this.rgbAuxIn,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed by using a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);

            #region Verify requirements related with NoCompression flag
            flagsInRgbOut = BitConverter.ToInt16(this.rgbOut, 2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R634, the value of Flags in RPC_HEADER_EXT is {0}", flagsInRgbOut);

            // Verify MS-OXCRPC requirement: R634
            // Call EcDoRpcExt2 whit RopLogon ROP command only return one RPC_HEADER_EXT.
            // R634 will be verified if flags in RPC_HEADER_EXT doesn't contain 0x0001, because NoCompression is contained in pulFlags when calling method EcDoRpcExt2.
            bool isVerifyR634 = (flagsInRgbOut & (short)RpcHeaderExtFlags.Compressed) != (short)RpcHeaderExtFlags.Compressed;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR634,
                634,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pulFlags] If pulFlags contains NoCompression (0x00000001), the server MUST NOT compress ROP response payload (rgbOut).");

            if (this.rgbAuxOut.Length != 0)
            {
                // Because NoCompression is contained in pulFlag when calling EcDoRpcExt2 method.
                // So R1304 will be verified if Flags in RPC_HEADER_EXT doesn't contain 0x0001.
                short flagsInRgbAuxOut = BitConverter.ToInt16(this.rgbAuxOut, 2);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1304, the value of Flags in RPC_HEADER_EXT is {0}.", flagsInRgbAuxOut);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1304
                bool isVerifyR1304 = (flagsInRgbAuxOut & (short)RpcHeaderExtFlags.Compressed) != (short)RpcHeaderExtFlags.Compressed;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1304,
                    1304,
                    @"[In EcDoRpcExt2 Method (opnum 11)] If pulFlags contains NoCompression (0x00000001), the server MUST NOT compress auxiliary payload (rgbAuxOut).");
            }
            #endregion
            #endregion

            #region Call EcDoRpcExt2 when pcxh is invalid
            this.pcxhInvalid = (IntPtr)ConstValues.InvalidPcxh;
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValueForInvalidCXH = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxhInvalid,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreNotEqual<uint>(0, this.returnValueForInvalidCXH, "EcDoRpcExt2 should not succeed by using an invalid CXH and '0' isn't expected to be returned. The returned value is {0}.", this.returnValueForInvalidCXH);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R392, the returned value of method EcDoRpcExt2 with parameter CXH returned from EcDoConnectEx is {0}, the returned value of method EcDoRpcExt2 with parameter CXH not returned from EcDoConnectEx is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R392
            // When the CXH is returned from EcDoConnectEx method, the call is successful. While the CXH is not returned from the EcDoConnectEx method, the call is failed. 
            // So if the code can reach here, the requirement is verified.
            bool isVerifyR392 = (this.returnValue == ResultSuccess) && (this.returnValueForInvalidCXH != ResultSuccess);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR392,
                392,
                @"[In Message Processing Events and Sequencing Rules] [EcDoRpcExt2] The method requires an active session context handle to be returned from the EcDoConnectEx method.");

            // Add the debug information 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R622, the returned value of method EcDoRpcExt2 with parameter CXH returned from EcDoConnectEx is {0}, the returned value of method EcDoRpcExt2 with parameter CXH not returned from EcDoConnectEx is {1}.", this.returnValue, this.returnValueForInvalidCXH);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R622
            // Since the context of R622 is similar with R392, if the R392 is verified, it means R622 is verified.
            bool isVerifyR622 = isVerifyR392;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR622,
                622,
                @"[In EcDoRpcExt2 Method (opnum 11)] This [method EcDoRpcExt2] call requires an active session context handle returned from the EcDoConnectEx method.");
            #endregion

            #region Call EcDoDisconnect with a valid CXH
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            // Because EcDummyRpc and EcDoRpcExt2 used the Session Context handle that is created by EcDoConnectEx method.
            // If the previous step of this case succeeds then whether the Session Context is persisted on the server will be verified.
            Site.CaptureRequirement(
                476,
                @"[In EcDoConnectEx Method (Opnum 10)] The Session Context is persisted on the server until the client disconnects by using the EcDoDisconnect method, as specified in section 3.1.4.3.");
        }

        /// <summary>
        /// This case tests calling EcDoConnectEx when ulFlags is 0x00000000 or 0x00000001.
        /// </summary>
        [TestCategory("MSOXCRPC"), TestMethod()]
        public void MSOXCRPC_S01_TC19_TestUlFlagsForEcDoConnectEx()
        {
            this.CheckTransport();
            uint flagsValue;

            #region Creates an RPC connection to the remote server use a user that not is administrator.
            string commonUserName = Common.GetConfigurationPropertyValue("NormalUserName", this.Site);
            string commonPassword = Common.GetConfigurationPropertyValue("NormalUserPassword", this.Site);

            this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, this.authenticationService, commonUserName, commonPassword);
            #endregion

            #region Call EcDoConnectEx method to establish a new Session Context with the server.
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            ushort[] rgwServerVersion;
            flagsValue = 0x00000000;
            string commonUserDN = Common.GetConfigurationPropertyValue("NormalUserEssdn", this.Site);
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                commonUserDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out rgwServerVersion,
                out this.rgwBestVersion,
                out this.picxr,
                flagsValue);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R489");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R489
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.ResultSuccess,
                this.returnValue,
                489,
                @"[In EcDoConnectEx Method (Opnum 10)] [ulFlags] [Value 0x00000000 means] requests connection without administrator privilege.");
            #endregion

            #region Call EcDoDisconnect to disconnect to server.
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Creates an RPC connection to the remote server use a user that is administrator.
            this.oxcrpcAdapter.InitializeRPC(this.authenticationLevel, this.authenticationService, this.userName, this.password);
            #endregion

            #region Call EcDoConnectEx method to establish a new Session Context with the server.
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            flagsValue = 0x00000001;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out rgwServerVersion,
                out this.rgwBestVersion,
                out this.picxr,
                flagsValue);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R490");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R490
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.ResultSuccess,
                this.returnValue,
                490,
                @"[In EcDoConnectEx Method (Opnum 10)] [ulFlags] [Value 0x00000001 means] Requests administrator behavior, which causes the server to check that the user has administrator privilege.");
            #endregion

            #region Call EcDoDisconnect to disconnect to server.
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "Call EcDoDisconnect method should succeed with a valid CXH and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
        }
        #endregion

        /// <summary>
        /// Initializes the test case before running it
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.rgbAuxIn = new byte[0];
            this.subFolderIds = new List<ulong>();
        }

        /// <summary>
        /// Clean up the test case after running it
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #region Private Methods

        #region Check whether the server supports all functionality

        /// <summary>
        /// Checks whether the server supports all functionality from previous server version level when supporting at one server version level.
        /// </summary>
        /// <param name="version">The server protocol version returned in rgwServerVersion field on method EcDoConnectEx</param>
        /// <param name="pcxh">The pointer points to the Context Handle returned on method EcDoConnectEx.</param>
        /// <returns>Returns true if supported. Otherwise, returns false</returns>
        private bool IsFunctionalitySupported(ushort[] version, IntPtr pcxh)
        {
            ushort[] normalizeVersion;
            bool isServerVersionSupport = false;
            AdapterHelper.ConvertVersion(version, out normalizeVersion);

            // Ignore the last (four) number: build minor number as this digit always means larger or equal to 0.
            // Transform the first three number to a long version in hexadecimal
            // For example: 14.3.312.233 will be transform to 000E00030138
            long longVersion = (normalizeVersion[0] * ConstValues.OffsetProductMajorVersion) +
                (normalizeVersion[1] * ConstValues.OffsetBuildMajorNumber) + normalizeVersion[2];

            // 600001A63 is calculated from 6.0.6755.0
            // The server version returned to the client should be greater than or equal to 6.0.6755.0 according to the Open Specification. So assert it fail once the returned version is less than 6.0.6755.0.
            if (longVersion < (long)ServerVersionValues.SupportBufferSizeField)
            {
                Site.Assert.Fail("The returned server version should be greater than or equal to 600001A63, which is calculated from 6.0.6755.0. Now the returned server version is {0}.", longVersion);
            }
            else
            {
                // Equal to or greater than 6.0.6755.0
                isServerVersionSupport = this.TryRopFastTransferSourceGetBuffer(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            // Equal to or greater than 8.0.295.0
            // 0x800000127 is calculated from 8.0.295.0
            if (longVersion >= (long)ServerVersionValues.SupportByteCountField)
            {
                isServerVersionSupport = this.TryRopReadStream(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            // Equal to or greater than 8.0.324.0
            // 0x800000144 is calculated from 8.0.324.0
            if (longVersion >= (long)ServerVersionValues.SupportOpenFlagsField)
            {
                isServerVersionSupport = this.TryRopLogon(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            // Equal to or greater than 8.0.358.0
            // 0x800000166 is calculated from 8.0.358.0
            if (longVersion >= (long)ServerVersionValues.SupportAsync)
            {
                isServerVersionSupport = this.TryEcDoAsyncConnectExandEcDoAsyncWaitEx(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            // Equal to or greater than 14.0.324.0
            // 0xE00000144 is calculated from 14.0.324.0
            if (longVersion >= (long)ServerVersionValues.SupportTableFlagsField)
            {
                isServerVersionSupport = this.TryRopGetContentsTable(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            // Equal to or greater than 14.0.616.0
            // 0xE00000268 is calculated from 14.0.616.0
            if (longVersion >= (long)ServerVersionValues.SupportImportDeleteFlagsField)
            {
                isServerVersionSupport = this.TryRopSynchronizationImportDeletes(pcxh);
                if (!isServerVersionSupport)
                {
                    return false;
                }
            }

            return isServerVersionSupport;
        }

        /// <summary>
        /// Verify whether server supports passing the sentinel value 0xBABE in the BufferSize field of a RopFastTransferSourceGetBuffer request.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports this ROP method with specific value passed in. True means yes, false means no.</returns>
        private bool TryRopFastTransferSourceGetBuffer(IntPtr pcxh)
        {
            // Maximum buffer size.
            const int MaxBufferSize = 0xFFFF;

            #region Logon
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, 0, (ulong)OpenFlags.UsePerMDBReplipMapping);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);

            // The element whose index is 0 indicates this ROP command response handle
            this.objHandle = this.responseSOHTable[0][logonResponse.OutputHandleIndex];
            #endregion

            #region CreateMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateMessage, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "RopCreateMessage should succeed and 0 is expected to be returned. The returned value is {0}.", createMessageResponse.ReturnValue);
            uint objCreateMessageHandle = this.responseSOHTable[0][createMessageResponse.OutputHandleIndex];
            #endregion

            #region OpenStream
            // OpenModeFlags is set to 0 means Create.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, objCreateMessageHandle, (ulong)StreamOpenModeFlags.Create);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenStreamResponse openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openStreamResponse.ReturnValue, "RopOpenStream should succeed and 0 is expected to be returned. The returned value is {0}.", openStreamResponse.ReturnValue);
            uint objOpenStreamHandle = this.responseSOHTable[0][openStreamResponse.OutputHandleIndex];
            #endregion

            #region RopWriteStream
            int writeCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("WriteStreamCount", Site));
            for (int counter = 0; counter < writeCount; counter++)
            {
                // auxInfo is not used for RopWriteSteam, so set it to 0
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopWriteStream, objOpenStreamHandle, 0);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref pcxh,
                    PulFlags.NoCompression | PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable);

                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
                RopWriteStreamResponse writeStreamResponse = (RopWriteStreamResponse)this.response;
                Site.Assert.AreEqual<uint>(0, writeStreamResponse.ReturnValue, "RopWriteStream should succeed and 0 is expected to be returned. The returned value is {0}.", writeStreamResponse.ReturnValue);
            }
            #endregion

            #region RopCommitStream
            // auxInfo is not used for RopCommitSteam, so set it to 0
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCommitStream, objOpenStreamHandle, 0);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopCommitStreamResponse commitStreamResponse = (RopCommitStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(0, commitStreamResponse.ReturnValue, "RopCommitStream should succeed and 0 is expected to be returned. The returned value is {0}.", commitStreamResponse.ReturnValue);
            #endregion

            #region SaveChangesMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSaveChangesMessage, objCreateMessageHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, saveChangesMessageResponse.ReturnValue, "RopSaveChangesMessage should succeed and 0 is expected to be returned. The returned value is {0}.", saveChangesMessageResponse.ReturnValue);
            #endregion

            #region OpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and 0 is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[0][openFolderResponse.OutputHandleIndex];
            #endregion

            #region FastTransferSourceCopyMessagesResponse
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopFastTransferSourceCopyMessages, this.objHandle, saveChangesMessageResponse.MessageId);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopFastTransferSourceCopyMessagesResponse fastTransferSourceCopyMessagesResponse = (RopFastTransferSourceCopyMessagesResponse)this.response;
            Site.Assert.AreEqual<uint>(0, fastTransferSourceCopyMessagesResponse.ReturnValue, "RopFastTransferSourceCopyMessages should succeed and 0 is expected to be returned. The returned value is {0}.", fastTransferSourceCopyMessagesResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[0][fastTransferSourceCopyMessagesResponse.OutputHandleIndex];
            #endregion

            #region RopFastTransferSourceGetBuffer
            // maximumBufferSize is set to 0xffff, BufferSize will be set to 0xBABE when composing ROP command.
            // Refer to the ComposeRopFastTransferSourceGetBufferRequest method.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopFastTransferSourceGetBuffer, this.objHandle, MaxBufferSize);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic | PulFlags.Chain,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopFastTransferSourceGetBufferResponse fastTransferSourceGetBufferResponse = (RopFastTransferSourceGetBufferResponse)this.response;
            Site.Assert.AreEqual<uint>(0, fastTransferSourceGetBufferResponse.ReturnValue, "RopFastTransferSourceGetBuffer should succeed and 0 is expected to be returned. The returned value is {0}.", fastTransferSourceGetBufferResponse.ReturnValue);
            #endregion
            return fastTransferSourceGetBufferResponse.ReturnValue == 0;
        }

        /// <summary>
        /// If server version equals to or greater than 8.0.295.0, server supports passing the sentinel value 0xBABE in the ByteCount field of a RopReadStream request.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports this ROP method with specific value passed in.</returns>
        private bool TryRopReadStream(IntPtr pcxh)
        {
            #region Logon
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, 0, (ulong)OpenFlags.UsePerMDBReplipMapping);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);

            // The element whose index is 0 indicates this ROP command response handle
            this.objHandle = this.responseSOHTable[0][logonResponse.OutputHandleIndex];
            #endregion

            #region CreateMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateMessage, this.objHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                createMessageResponse.ReturnValue,
                "If RopCreateMessage succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                createMessageResponse.ReturnValue);
            uint messageHandle = this.responseSOHTable[0][createMessageResponse.OutputHandleIndex];
            #endregion

            #region OpenStream
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, messageHandle, (ulong)StreamOpenModeFlags.Create);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenStreamResponse openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openStreamResponse.ReturnValue,
               "If RopOpenStream succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                openStreamResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[0][openStreamResponse.OutputHandleIndex];
            #endregion

            #region RopWriteStream
            RopWriteStreamResponse writeStreamResponse;
            int writeCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("WriteStreamCount", Site));
            for (int counter = 0; counter < writeCount; counter++)
            {
                // auxInfo is not used for RopWriteSteam, so set it to 0
                this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopWriteStream, this.objHandle, 0);
                this.pcbOut = ConstValues.ValidpcbOut;
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.responseSOHTable = new List<List<uint>>();

                this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                    ref pcxh,
                    PulFlags.NoCompression | PulFlags.NoXorMagic,
                    this.rgbIn,
                    ref this.pcbOut,
                    null,
                    ref this.pcbAuxOut,
                    out this.response,
                    ref this.responseSOHTable);

                Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
                writeStreamResponse = (RopWriteStreamResponse)this.response;

                Site.Assert.AreEqual<uint>(
                    0x00000000,
                    writeStreamResponse.ReturnValue,
                    "If RopWriteStream succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                    writeStreamResponse.ReturnValue);
            }
            #endregion

            #region RopCommitStream
            // auxInfo is not used for RopCommitSteam, so set it to 0
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCommitStream, this.objHandle, 0);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopCommitStreamResponse commitStreamResponse = (RopCommitStreamResponse)this.response;

            Site.Assert.AreEqual<uint>(
                0x00000000,
                commitStreamResponse.ReturnValue,
                "If RopCommitStream succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                commitStreamResponse.ReturnValue);
            #endregion

            #region OpenStream
            // OpenModeFlags is set to 0 means ReadOnly.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenStream, messageHandle, 0);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            openStreamResponse = (RopOpenStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openStreamResponse.ReturnValue,
                "If RopOpenStream succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                openStreamResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[0][openStreamResponse.OutputHandleIndex];
            #endregion

            #region ReadStream
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopReadStream, this.objHandle, ConstValues.MaximumByteCountIndicator);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic | PulFlags.Chain,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopReadStreamResponse readStreamResponse = (RopReadStreamResponse)this.response;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                readStreamResponse.ReturnValue,
                "If RopReadStream succeeds, the ReturnValue of its response is 0(success). Now the returned value is {0}.",
                readStreamResponse.ReturnValue);
            #endregion
            return readStreamResponse.ReturnValue == 0;
        }

        /// <summary>
        /// Verify whether server supports the flag CLI_WITH_PER_MDB_FIX in the OpenFlags field of a RopLogon request.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports this ROP method with specific value passed in. True means yes, false means no.</returns>
        private bool TryRopLogon(IntPtr pcxh)
        {
            // CLI_WITH_PER_MDB_FIX in the OpenFlags field
            #region Logon
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, 0, (ulong)OpenFlags.UsePerMDBReplipMapping);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);

            // The element whose index is 0 indicates this ROP command response handle
            this.objHandle = this.responseSOHTable[0][logonResponse.OutputHandleIndex];
            #endregion

            return logonResponse.ReturnValue == 0;
        }

        /// <summary>
        /// Verify whether server supports the EcDoAsyncConnectEx and EcDoAsyncWaitEx RPC function calls.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports RPC methods EcDoAsyncConnectEx and EcDoAsyncWaitEx. True means yes, false means no.</returns>
        private bool TryEcDoAsyncConnectExandEcDoAsyncWaitEx(IntPtr pcxh)
        {
            // Supports the EcDoAsyncConnectEx and EcDoAsyncWaitEx
            #region Call EcDoAsyncConnectEx
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncConnectEx(pcxh, ref this.pacxh);
            if (this.returnValue != 0)
            {
                return false;
            }
            #endregion

            #region Call EcDoAsyncWaitEx
            bool isNotificationPending;
            this.returnValue = this.oxcrpcAdapter.EcDoAsyncWaitEx(this.pacxh, out isNotificationPending);
            #endregion call EcDoAsyncWaitEx

            return this.returnValue == 0;
        }

        /// <summary>
        /// Verify whether server supports passing the flag ConversationMembers (0x80) in the TableFlags field of a RopGetContentsTable request.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports this ROP method with specific value passed in. True means yes, false means no.</returns>
        private bool TryRopGetContentsTable(IntPtr pcxh)
        {
            #region Logon
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, 0, (ulong)OpenFlags.UsePerMDBReplipMapping);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);

            // The element whose index is 0 indicates this ROP command response handle
            this.objHandle = this.responseSOHTable[0][logonResponse.OutputHandleIndex];
            #endregion

            #region OpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, this.objHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and 0 is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            this.objHandle = this.responseSOHTable[0][openFolderResponse.OutputHandleIndex];
            #endregion

            #region GetContentsTable
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopGetContentsTable, this.objHandle, ConstValues.ConversationMemberTableFlag);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
            ref pcxh,
            PulFlags.NoCompression | PulFlags.NoXorMagic,
            this.rgbIn,
            ref this.pcbOut,
            null,
            ref this.pcbAuxOut,
            out this.response,
            ref this.responseSOHTable);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopGetContentsTableResponse getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(0, getContentsTableResponse.ReturnValue, "RopGetContentsTable should succeed and 0 is expected to be returned. The returned value is {0}.", getContentsTableResponse.ReturnValue);
            #endregion

            return getContentsTableResponse.ReturnValue == 0;
        }

        /// <summary>
        /// Verify whether server supports passing the flag HardDelete (0x02) in the Flags field of a RopSynchronizationImportDeletes request.
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a session context handle.</param>
        /// <returns>Whether the server supports this ROP method with specific value passed in. True means yes, false means no.</returns>
        private bool TryRopSynchronizationImportDeletes(IntPtr pcxh)
        {
            #region Logon to mailbox
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, 0, (ulong)OpenFlags.UsePerMDBReplipMapping);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "RopLogon should succeed and 0 is expected to be returned. The returned value is {0}.", logonResponse.ReturnValue);

            // The element whose index is 0 indicates this ROP command response handle
            this.objHandle = this.responseSOHTable[0][logonResponse.OutputHandleIndex];
            #endregion

            #region OpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, this.objHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openFolderResponse.ReturnValue, "RopOpenFolder should succeed and 0 is expected to be returned. The returned value is {0}.", openFolderResponse.ReturnValue);
            uint openObjHandle = this.responseSOHTable[0][openFolderResponse.OutputHandleIndex];
            #endregion

            #region Configure a synchronization upload context
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSynchronizationOpenCollector, openObjHandle, 0);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopSynchronizationOpenCollectorResponse openCollectorResponse = (RopSynchronizationOpenCollectorResponse)this.response;
            Site.Assert.AreEqual<uint>(0, openCollectorResponse.ReturnValue, "RopSynchronizationOpenCollector should succeed and 0 is expected to be returned. The returned value is {0}.", openCollectorResponse.ReturnValue);
            uint synchronizationUploadContextHandle = this.responseSOHTable[0][openCollectorResponse.OutputHandleIndex];
            #endregion

            #region RopCreateMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopCreateMessage, openObjHandle, logonResponse.FolderIds[(int)FolderIds.Inbox]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "RopCreateMessage should succeed and 0 is expected to be returned. The returned value is {0}.", createMessageResponse.ReturnValue);
            uint objCreateMessageHandle = this.responseSOHTable[0][createMessageResponse.OutputHandleIndex];
            #endregion

            #region RopSaveChangesMessage
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSaveChangesMessage, objCreateMessageHandle, logonResponse.FolderIds[(int)FolderIds.InterpersonalMessage]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(0, saveChangesMessageResponse.ReturnValue, "RopSaveChangesMessage should succeed and 0 is expected to be returned. The returned value is {0}.", saveChangesMessageResponse.ReturnValue);
            #endregion

            #region RopSynchronizationImportDeletes

            #region RopLongTermIdFromIdRequest
            // Call RopLongTermIdFromIdRequest to convert the short-term ID into a long-term ID for RopSynchronizationImportDeletes.
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLongTermIdFromId, this.objHandle, saveChangesMessageResponse.MessageId);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopLongTermIdFromIdResponse longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.response;
            Site.Assert.AreEqual<uint>(0, longTermIdFromIdResponse.ReturnValue, "RopLongTermIdFromId should succeed and 0 is expected to be returned. The returned value is {0}.", longTermIdFromIdResponse.ReturnValue);

            // PropertyValues for RopSynchronizationImportDeletes.
            byte[] importdeletesPropertyValues = new byte[sizeof(int) + sizeof(short) + ConstValues.GidLength];
            PropertyTag[] tagArray = new PropertyTag[1];

            TaggedPropertyValue propertyValue = new TaggedPropertyValue();
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = 0x0000,
                PropertyType = 0x1102
            };

            int index = 0;
            Array.Copy(BitConverter.GetBytes(tagArray.Length), 0, importdeletesPropertyValues, 0, sizeof(int));
            index += sizeof(int);
            Array.Copy(BitConverter.GetBytes(22), 0, importdeletesPropertyValues, index, sizeof(short));
            index += sizeof(short);
            byte[] longTermByte = new byte[longTermIdFromIdResponse.LongTermId.DatabaseGuid.Length + longTermIdFromIdResponse.LongTermId.GlobalCounter.Length];
            Array.Copy(longTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, longTermByte, 0, longTermIdFromIdResponse.LongTermId.DatabaseGuid.Length);
            Array.Copy(longTermIdFromIdResponse.LongTermId.GlobalCounter, 0, longTermByte, longTermIdFromIdResponse.LongTermId.DatabaseGuid.Length, longTermIdFromIdResponse.LongTermId.GlobalCounter.Length);

            Array.Copy(longTermByte, 0, importdeletesPropertyValues, index, longTermByte.Length);
            propertyValue.PropertyTag = propertyTag;
            propertyValue.Value = importdeletesPropertyValues;
            #endregion

            // HardDelete (0x02) is written in method ComposeRopSynchronizationImportDeletes
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopSynchronizationImportDeletes, synchronizationUploadContextHandle, propertyValue);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed and 0 is expected to be returned. The returned value is {0}.", this.returnValue);
            RopSynchronizationImportDeletesResponse synchronizationImportDeletesResponse = (RopSynchronizationImportDeletesResponse)this.response;
            Site.Assert.AreEqual<uint>(0, synchronizationImportDeletesResponse.ReturnValue, "RopSynchronizationImportDeletes should succeed and 0 is expected to be returned. The returned value is {0}.", synchronizationImportDeletesResponse.ReturnValue);
            #endregion

            return this.returnValue == 0;
        }
        #endregion

        /// <summary>
        /// Verify whether server can add additional data.
        /// </summary>
        /// <param name="payloadCount">The count of payload that ROP response contains.</param>
        private void ServerAddAdditionalData(uint payloadCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R962, the payload count is: {0}", payloadCount);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R962
            // Server can add additional data indicates the rgbOut has two payloads, so payloadCount > 1. If the code can reach here indicates rgbOut response buffer with its own RPC_HEADER_EXT header, this requirement is verified.
            Site.CaptureRequirementIfIsTrue(
                payloadCount > 1,
                962,
                @"[In Extended Buffer Packing] However, when the server finishes processing a RopQueryRows ROP ([MS-OXCROPS] section 2.2.5.4), RopReadStream ROP ([MS-OXCROPS] section 2.2.9.2), or RopFastTransferSourceGetBuffer ROP ([MS-OXCROPS] section 2.2.12.3) from the rgbIn request buffer and it [actual ROP request] was the last ROP command in the request and the client has requested packing through the Chain flag and there is residual room in the rgbOut response buffer, the server can add additional data to the rgbOut parameter response, each with its own RPC_HEADER_EXT header.");
        }

        /// <summary>
        /// Parses a response with multiple ROPs. In this scenario, the response is designed as only containing two ROPs: RopSetColumn and RopQueryRows.
        /// </summary>
        /// <param name="rgbOutput">The raw data that contains the ROP response payload</param>
        /// <returns>The ROP response list which contains RopSetColumnResponse and RopQueryRawResponse</returns>
        private List<IDeserializable> ParseMultipleRopsResponse(byte[] rgbOutput)
        {
            int parseByteLength = 0;
            List<IDeserializable> multipleRopsResponse = new List<IDeserializable>();
            RPC_HEADER_EXT rpcHeader = new RPC_HEADER_EXT
            {
                Version = BitConverter.ToUInt16(rgbOutput, parseByteLength)
            };

            // Parse RPC_HEADER_EXT structure
            parseByteLength += sizeof(short);
            rpcHeader.Flags = BitConverter.ToUInt16(rgbOutput, parseByteLength);
            parseByteLength += sizeof(short);
            rpcHeader.Size = BitConverter.ToUInt16(rgbOutput, parseByteLength);
            parseByteLength += sizeof(short);
            rpcHeader.SizeActual = BitConverter.ToUInt16(rgbOutput, parseByteLength);
            parseByteLength += sizeof(short);

            // Passed 2 bytes which is size of ROP response.
            parseByteLength += sizeof(short);

            // Parse RopSetColumns response
            RopSetColumnsResponse setColumnsResponse = new RopSetColumnsResponse();
            parseByteLength += setColumnsResponse.Deserialize(rgbOutput, parseByteLength);

            // Parse RopQueryRows response
            RopQueryRowsResponse queryRowsResponse = new RopQueryRowsResponse();
            parseByteLength += queryRowsResponse.Deserialize(rgbOutput, parseByteLength);

            multipleRopsResponse.Add(setColumnsResponse);
            multipleRopsResponse.Add(queryRowsResponse);

            return multipleRopsResponse;
        }

        /// <summary>
        /// It is used to test the reserved fields in the rgbAuxIn on method EcDoConnectEx.
        /// </summary>
        /// <param name="rgbAuxInEnum">Enum of the blocks sent in the rgbAuxIn on method EcDoConnectEx</param>
        /// <param name="payload">A byte array contains an auxiliary payload buffer.</param>
        /// <returns>The return value of the EcDoConnectEx method</returns>
        private uint SendAuxiliaryPayloadBufferInEcDoConnectEx(RgbAuxInEnum rgbAuxInEnum, byte[] payload)
        {
            uint returnValueOfDisconnect = 0;
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(rgbAuxInEnum, payload);
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed for test of reserved fields and '0' is expected to be returned. The returned value is {0}.", this.returnValue);

            returnValueOfDisconnect = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, returnValueOfDisconnect, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", returnValueOfDisconnect);
            return this.returnValue;
        }

        /// <summary>
        /// It is used to test the Reserved fields in the rgbAuxIn on method EcDoRpcExt2
        /// </summary>
        /// <param name="rgbAuxInEnum">Enum of the blocks sent in the rgbAuxIn on method EcDoRpcExt2</param>
        /// <param name="payload">A byte array contains an auxiliary payload buffer.</param>
        /// <returns>The return value of the EcDoRpcExt2 method</returns>
        private uint SendAuxiliaryPayloadBufferInEcDoRpcExt2(RgbAuxInEnum rgbAuxInEnum, byte[] payload)
        {
            #region Call EcDoConnectEx to establish a new Session Context with the server.
            uint returnValueOfDisconnect = 0;
            byte[] sessionInfo = AdapterHelper.Compose_AUX_PERF_SESSIONINFO(ReserveDefault);
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(RgbAuxInEnum.AUX_PERF_SESSIONINFO, sessionInfo);

            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                this.rgbAuxIn,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed and send CXH to EcDoRpcExt2. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoRpcExt2 to pass generic remote operation (ROP) commands to the server.
            this.rgbAuxIn = AdapterHelper.ComposeRgbAuxIn(rgbAuxInEnum, payload);

            // Parameter inObjHandle is no use for RopLogon command, so set it to unUsedInfo
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                this.rgbAuxIn,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoRpcExt2 should succeed for test of the Reserved fields in the rgbAuxIn. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Call EcDoDisconnect to close a Session Context with the server
            returnValueOfDisconnect = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, returnValueOfDisconnect, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion
            return this.returnValue;
        }

        /// <summary>
        /// Receive push notification on the specified port.
        /// </summary>
        /// <param name="addressFamily">Specifies which IP family to use.</param>
        /// <returns>The opaque data received from server.</returns>
        private string GetPushNotification(Add_Families addressFamily)
        {
            string opaque = null;
            int port = int.Parse(Common.GetConfigurationPropertyValue("NotificationPort", this.Site));

            using (UdpClient udpClient = new UdpClient(port, addressFamily == Add_Families.AF_INET ? System.Net.Sockets.AddressFamily.InterNetwork : System.Net.Sockets.AddressFamily.InterNetworkV6))
            {
                udpClient.Client.ReceiveTimeout = int.Parse(Common.GetConfigurationPropertyValue("ReceiveTimeout", this.Site));

                IPEndPoint remote = new IPEndPoint(IPAddress.Any, 0);
                byte[] data;
                try
                {
                    data = udpClient.Receive(ref remote);
                    opaque = System.Text.Encoding.ASCII.GetString(data).Replace("\0", string.Empty);
                    return opaque;
                }
                catch (SocketException exception)
                {
                    Site.Assert.Fail("Failed to receive the push notification. The error message is {0}.", exception.Message);
                    return string.Empty;
                }
            }
        }
        #endregion
    }
}