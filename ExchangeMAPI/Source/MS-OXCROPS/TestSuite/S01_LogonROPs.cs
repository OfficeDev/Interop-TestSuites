namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is designed to verify the response buffer formats of Logon ROPs. 
    /// </summary>
    [TestClass]
    public class S01_LogonROPs : TestSuiteBase
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
        /// This method tests the failure response buffer of RopLogon.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC01_TestLogonFailed()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            RopLogonRequest logonRequest;
            RopLogonResponse logonResponse;

            logonRequest.RopId = (byte)RopId.RopLogon;
            logonRequest.LogonId = TestSuiteBase.LogonId;

            // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            logonRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex0;

            logonRequest.LogonFlags = (byte)LogonFlags.Private;
            logonRequest.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;
            logonRequest.StoreState = (uint)StoreState.None;

            // Set Essdn to the content of user DN, which specifies it will log on to the mail box of user represented by the user DN.
            logonRequest.Essdn = Encoding.ASCII.GetBytes(TestSuiteBase.WrongUserDN + "\0");

            // Set EssdnSize to the byte count of user DN, which specifies the size of the Essdn field.
            logonRequest.EssdnSize = (ushort)logonRequest.Essdn.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopLogon request.");

            // Send the RopLogon request to the server.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                logonRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            logonResponse = (RopLogonResponse)response;

            // Verify the response is a failure response.
            bool isFailureResponse = (logonResponse.ReturnValue != TestSuiteBase.SuccessReturnValue) && (logonResponse.ReturnValue != MS_OXCROPSAdapter.WrongServer);
            Site.Assert.IsTrue(isFailureResponse, "For this response, this field is set to a value other than 0x00000000 or 0x00000478.");
        }

        /// <summary>
        /// This method tests the Redirect response buffer of RopLogon.
        /// This test case depends on the second SUT. If the second SUT is not present, this test case cannot be executed.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC02_TestLogonRedirect()
        {
            this.CheckTransportIsSupported();

            // According to the Open Specification MS-OXCMAPIHTTP, AutoDiscover must be used when transport sequence is set to mapi_http no mater 
            // whether AutoDiscover is enabled or disabled in the common configuration file. In this case, AutoDiscover must be disabled which doesn't 
            // match the basic prerequisite of transport sequence mapi_http. So this case can't be run under transport sequence mapi_http.
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() != "mapi_http")
            {
                // Check whether the environment supported public folder.
                if (bool.Parse(Common.GetConfigurationPropertyValue("IsPublicFolderSupported", this.Site)))
                {
                    // Check whether the environment supported redirect server.
                    if (!string.IsNullOrEmpty(Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site)))
                    {
                        // Get parameters values from configure file, which will be used in the following RpcConnect method.
                        string server = Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site);
                        string userDN = Common.GetConfigurationPropertyValue("UserEssdn", this.Site) + '\0';
                        string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                        string userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
                        string password = Common.GetConfigurationPropertyValue("PassWord", this.Site);
                        cropsAdapter.SetAutoRedirect(false);

                        // Connect to the server.
                        bool isConnected = cropsAdapter.RpcConnect(server, ConnectionType.PublicFolderServer, userDN, domain, userName, password);

                        // Check the connect status.
                        Site.Assert.IsTrue(
                            isConnected,
                            "RPC connect to {0} fails.",
                            server);

                        RopLogonRequest logonRequest;

                        logonRequest.RopId = (byte)RopId.RopLogon;
                        logonRequest.LogonId = TestSuiteBase.LogonId;

                        // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
                        // where the handle for the output Server object will be stored.
                        logonRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex0;

                        logonRequest.StoreState = (uint)StoreState.None;

                        // Set other parameters for logon type of PublicFolder.
                        logonRequest.LogonFlags = (byte)LogonFlags.PublicFolder;
                        logonRequest.OpenFlags = (uint)OpenFlags.Public;

                        // Set EssdnSize to 0, which specifies the size of the Essdn field.
                        logonRequest.EssdnSize = 0;

                        // Initialize the Essdn to null.
                        logonRequest.Essdn = null;

                        // Send the RopLogon request and get the response.
                        cropsAdapter.ProcessSingleRop(
                            logonRequest,
                            this.inputObjHandle,
                            ref this.response,
                            ref this.rawData,
                            RopResponseType.RedirectResponse);
                        cropsAdapter.SetAutoRedirect(true);
                    }
                    else
                    {
                        Site.Assert.Inconclusive("This case runs only when the second system under test exists.");
                    }
                }
                else
                {
                    Site.Assert.Inconclusive("This case runs only when the system under test supports public folder logon.");
                }
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the Autodiscover is disabled and transport sequence is not mapi_http since mapi_http requires Autodicover is enabled.");
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetStoreState, RopSetReceiveFolder, RopGetReceiveFolder and RopGetReceiveFolderTable.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC03_TestRopReceiveFolder()
        {
            this.CheckTransportIsSupported();

            #region Local variables

            // Define the index mark of that specified message class associates FoldId in RopLogonResponse.FolderIds.
            int indexOldFolderId = -1;

            // Also define the index variable for RopSetReceiveFolderId.
            int indexNewFolderId = -1;

            #endregion

            // Step 1: Send a RopGetStoreState request and verify the RopGetStoreState success response.
            #region RopGetStoreState success response

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopGetStoreStateRequest getStoreStateRequest;
            RopGetStoreStateResponse getStoreStateResponse;

            getStoreStateRequest.RopId = (byte)RopId.RopGetStoreState;
            getStoreStateRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getStoreStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            if (Common.IsRequirementEnabled(312601, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopGetStoreState request.");

                // Send the RopGetStoreState request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getStoreStateRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getStoreStateResponse = (RopGetStoreStateResponse)response;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R312601");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R312601
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getStoreStateResponse.ReturnValue,
                    312601,
                    @"[In Appendix A: Product Behavior] Implementation does implement the RopGetStoreState remote operation (ROP). (Exchange 2007 follow this behavior.)");

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getStoreStateResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");
            }

            #endregion

            // Step 2: Send a RopGetReceiveFolder request and verify the RopGetReceiveFolder success response.
            #region RopGetReceiveFolder success response

            RopGetReceiveFolderRequest getReceiveFolderRequest;
            RopGetReceiveFolderResponse getReceiveFolderResponse;

            getReceiveFolderRequest.RopId = (byte)RopId.RopGetReceiveFolder;
            getReceiveFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getReceiveFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MessageClass, which specifies which message class to set the receive folder for.
            getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(TestSuiteBase.MessageClassForReceiveFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetReceiveFolder request.");

            // Send the RopGetReceiveFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getReceiveFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getReceiveFolderResponse = (RopGetReceiveFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getReceiveFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Verify the Receive folder id is valid.
            bool isFolderIdValid = false;
            for (int i = 0; i < logonResponse.FolderIds.Length; i++)
            {
                if (logonResponse.FolderIds[i] == getReceiveFolderResponse.FolderId)
                {
                    indexOldFolderId = i;
                    isFolderIdValid = true;
                }
            }

            Site.Assert.AreEqual<bool>(
                true,
                isFolderIdValid,
                "The Receive folder MUST be a folder within the user's mailbox.");

            #endregion

            // Step 3: Send a RopSetReceiveFolder request and verify the RopSetReceiveFolder success response.
            #region RopSetReceiveFolder success response

            RopSetReceiveFolderRequest setReceiveFolderRequest;
            RopSetReceiveFolderResponse setReceiveFolderResponse;

            setReceiveFolderRequest.RopId = (byte)RopId.RopSetReceiveFolder;
            setReceiveFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            setReceiveFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set FolderId to different folder of this mailbox.
            indexNewFolderId = indexOldFolderId >= (logonResponse.FolderIds.Length - 1) ? indexOldFolderId - 1 : indexOldFolderId + 1;
            setReceiveFolderRequest.FolderId = logonResponse.FolderIds[indexNewFolderId];

            // Set MessageClass, which specifies which message class to set the receive folder for.
            setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(TestSuiteBase.MessageClassForReceiveFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSetReceiveFolder request.");

            // Send the RopSetReceiveFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setReceiveFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setReceiveFolderResponse = (RopSetReceiveFolderResponse)response;

            #endregion

            // Step 4: Send a RopGetReceiveFolder request and verify the RopGetReceiveFolder success response.
            #region RopGetReceiveFolder success response

            getReceiveFolderRequest.RopId = (byte)RopId.RopGetReceiveFolder;
            getReceiveFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getReceiveFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set MessageClass, which specifies which message class to set the receive folder for.
            getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(TestSuiteBase.MessageClassForReceiveFolder + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetReceiveFolder request.");

            // Send the RopGetReceiveFolder request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getReceiveFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getReceiveFolderResponse = (RopGetReceiveFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getReceiveFolderResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Verify the Receive folder id is valid.
            isFolderIdValid = false;
            foreach (ulong folderId in logonResponse.FolderIds)
            {
                if (folderId == getReceiveFolderResponse.FolderId)
                {
                    isFolderIdValid = true;
                }
            }

            Site.Assert.AreEqual<bool>(
                true,
                isFolderIdValid,
                "The Receive folder MUST be a folder within the user's mailbox.");

            #endregion

            // Verification for RopSetReceiveFolder.
            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setReceiveFolderResponse.ReturnValue,
                "If RopSetReceiveFolder ROP succeeds, the ReturnValue of its response should be 0 (success)");

            Site.Assert.AreNotEqual<ulong>(
               logonResponse.FolderIds[indexOldFolderId],
               logonResponse.FolderIds[indexNewFolderId],
               "If RopSetReceiveFolder ROP succeeds, the id of receive folder should be different with the original one.");

            // Step 5: Send a GetReceiveFolderTable request and verify the GetReceiveFolderTable success response.
            #region GetReceiveFolderTable success response

            RopGetReceiveFolderTableRequest getReceiveFolderTableRequest;
            RopGetReceiveFolderTableResponse getReceiveFolderTableResponse;

            getReceiveFolderTableRequest.RopId = (byte)RopId.RopGetReceiveFolderTable;
            getReceiveFolderTableRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getReceiveFolderTableRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetReceiveFolderTable request.");

            // Send the GetReceiveFolderTable request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getReceiveFolderTableRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getReceiveFolderTableResponse = (RopGetReceiveFolderTableResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getReceiveFolderTableResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Check whether the environment supported public folder.
            if (bool.Parse(Common.GetConfigurationPropertyValue("IsPublicFolderSupported", this.Site)))
            {
                // Reconnect to public folder.
                bool ret = this.cropsAdapter.RpcDisconnect();
                this.Site.Assert.IsTrue(ret, "Rpc disconnect should be success.");
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PublicFolderServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));

                // Step 6: Send a RopGetStoreState request and verify the RopGetStoreState failure response.
                #region RopGetStoreState failure response

                // Log on to the public folder.
                logonResponse = this.Logon(LogonType.PublicFolder, this.userDN, out this.inputObjHandle);

                // Set InputHandleIndex to 0x1, which is an invalid index and will lead to a failure response.
                getStoreStateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

                if (Common.IsRequirementEnabled(312601, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopGetStoreState request.");

                    // Send the RopGetStoreState request to the server and verify the failure response.
                    this.responseSOHs = cropsAdapter.ProcessSingleRop(
                        getStoreStateRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.FailureResponse);
                    getStoreStateResponse = (RopGetStoreStateResponse)response;

                    Site.Assert.AreNotEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        getStoreStateResponse.ReturnValue,
                        "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");
                }

                #endregion

                // Step 7: Send a RopGetReceiveFolder request and verify the RopGetReceiveFolder failure response.
                #region RopGetReceiveFolder failure response

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopGetReceiveFolder request.");

                // Send the RopGetReceiveFolder request to the server and verify the failure response.
                // The server verifies that the operation is being performed against a public folders logon, which failed the operation with the ReturnValue field 0x80040102.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getReceiveFolderRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getReceiveFolderResponse = (RopGetReceiveFolderResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getReceiveFolderResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion

                // Step 8: Send a RopGetReceiveFolderTable request and verify the RopGetReceiveFolderTable failure response.
                #region RopGetReceiveFolderTable failure response

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopGetReceiveFolderTable request.");

                // Send the RopGetReceiveFolderTable request to the server and verify the failure response.
                // The server verifies that the operation is being performed against a public folders logon, which failed the operation with the ReturnValue field 0x80040102.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getReceiveFolderTableRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getReceiveFolderTableResponse = (RopGetReceiveFolderTableResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getReceiveFolderTableResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopGetOwningServers.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC04_TestRopGetOwningServers()
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

                // Log on to the public folder, for this operation SHOULD be issued against a public folders logon.
                RopLogonResponse logonResponse = Logon(LogonType.PublicFolder, this.userDN, out inputObjHandle);

                // Step 1: Open the second folder, in which a public folder will be created under root folder in the following code.
                #region Open the second folder

                RopOpenFolderRequest openFolderRequest;
                RopOpenFolderResponse openFolderResponse;

                openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

                openFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
                // where the handle for the output Server object will be stored.
                openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                // Set FolderId to the second folder.
                openFolderRequest.FolderId = logonResponse.FolderIds[1];

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
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                // Get the folder handle which is opened.
                uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

                #endregion

                // Step 2: Create a public folder under the root folder.
                #region Create folder

                RopCreateFolderRequest createFolderRequest;
                RopCreateFolderResponse createFolderResponse;

                createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
                createFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
                // where the handle for the output Server object will be stored.
                createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

                // Set UseUnicodeStrings to 0x0, which specifies the DisplayName and Comment are not specified in Unicode.
                createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

                // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
                createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

                // Set Reserved to 0x0, this field is reserved and MUST be set to 0.
                createFolderRequest.Reserved = TestSuiteBase.Reserved;

                // Set DisplayName, which specifies the name of the created folder.
                createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Common.GenerateResourceName(this.Site, "PublicFolderNoneGhosted") + "\0");

                // Set Comment, which specifies the folder comment that is associated with the created folder.
                createFolderRequest.Comment = Encoding.ASCII.GetBytes(Common.GenerateResourceName(this.Site, "PublicFolderNoneGhosted") + "\0");

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

                // Get the folder id of the created folder.
                ulong folderId = createFolderResponse.FolderId;

                #endregion

                // Step 3: Send the RopLogon request and verify the success response.
                #region Successful response

                RopGetOwningServersRequest getOwningServersRequest;
                RopGetOwningServersResponse getOwningServersResponse;

                getOwningServersRequest.RopId = (byte)RopId.RopGetOwningServers;
                getOwningServersRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                getOwningServersRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
                getOwningServersRequest.FolderId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetOwningServers request.");

                // Send the RopGetOwningServers request to the server and verify the success response. 
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getOwningServersRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getOwningServersResponse = (RopGetOwningServersResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getOwningServersResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step4: Send a RopDeleteFolder request to the server.
                #region RopDeleteFolder Response

                RopDeleteFolderRequest deleteFolderRequest;

                deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
                deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                deleteFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;

                // Set FolderId to targetFolderId, this folder is to be deleted.
                deleteFolderRequest.FolderId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopDeleteFolder request.");

                // Send a RopDeleteFolder request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    deleteFolderRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    deleteFolderResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 5: Send the RopLogon request and verify the failure response.
                #region Failure response

                // Set FolderId to 0x1, which does not exist in public folder database and will lead to a failure response.
                getOwningServersRequest.FolderId = TestSuiteBase.WrongFolderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetOwningServers request.");

                // Send the RopGetOwningServers request to the server and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getOwningServersRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getOwningServersResponse = (RopGetOwningServersResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getOwningServersResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the first system supports public folder logon.");
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopPublicFolderIsGhosted.
        /// This test case depends on the second SUT. If the second SUT is not present, some steps (6-9) of this test case cannot be executed.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC05_TestRopPublicFolderIsGhosted()
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

                // Step 1: Open the second folder, in which a public folder will be created under root folder in the following code.
                #region Open the second folder
                RopLogonResponse logonResponse = Logon(LogonType.PublicFolder, this.userDN, out inputObjHandle);
                RopOpenFolderRequest openFolderRequest;
                RopOpenFolderResponse openFolderResponse;

                openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

                openFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
                // where the handle for the output Server object will be stored.
                openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                // Set FolderId to the second folder.
                openFolderRequest.FolderId = logonResponse.FolderIds[1];

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
                    "if ROP succeeds, the ReturnValue of its response is 0(success)");

                // Get the folder handle which is opened.
                uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];

                #endregion

                // Step 2: Create a none-ghosted public folder under the root folder.
                #region Create folder

                RopCreateFolderRequest createFolderRequest;
                RopCreateFolderResponse createFolderResponse;

                createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
                createFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set OutputHandleIndex to 0x1, which specifies the location in the Server object handle table
                // where the handle for the output Server object will be stored.
                createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                createFolderRequest.FolderType = (byte)FolderType.Genericfolder;

                // Set UseUnicodeStrings to 0x0, which specifies the DisplayName and Comment are not specified in Unicode.
                createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);

                // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
                createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

                // Set Reserved to 0x0, this field is reserved and MUST be set to 0.
                createFolderRequest.Reserved = TestSuiteBase.Reserved;

                // Set DisplayName, which specifies the name of the created folder.
                createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Common.GenerateResourceName(this.Site, "PublicFolderNoneGhosted") + "\0");

                // Set Comment, which specifies the folder comment that is associated with the created folder.
                createFolderRequest.Comment = Encoding.ASCII.GetBytes(Common.GenerateResourceName(this.Site, "PublicFolderNoneGhosted") + "\0");

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

                // Get the folder id of the created folder.
                ulong folderId = createFolderResponse.FolderId;

                #endregion

                // Step 3: Send RopPublicFolderIsGhosted request to the server and verify RopPublicFolderIsGhosted success response with none-Ghosted folder.
                #region RopPublicFolderIsGhosted success response with none-Ghosted folder

                RopPublicFolderIsGhostedRequest publicFolderIsGhostedRequest;
                RopPublicFolderIsGhostedResponse publicFolderIsGhostedResponse;

                publicFolderIsGhostedRequest.RopId = (byte)RopId.RopPublicFolderIsGhosted;
                publicFolderIsGhostedRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                publicFolderIsGhostedRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set the FolderId to that got in Step 2.
                publicFolderIsGhostedRequest.FolderId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopPublicFolderIsGhosted request.");

                // Send the RopPublicFolderIsGhosted request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    publicFolderIsGhostedRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    publicFolderIsGhostedResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                #region Verify R4630, R313 and R4631
                Site.Assert.AreEqual<uint>(
                    0x0,
                    publicFolderIsGhostedResponse.IsGhosted,
                    "If the test case is opening a non-ghosted public folder, IsGhosted should be 0");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4630");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4630
                // This ServersCount is not present when ServersCount is null.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    publicFolderIsGhostedResponse.ServersCount,
                    4630,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ServersCount (2 bytes): This field[ServersCount (2 bytes)] is not present if IsGhosted is zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R313");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R313
                // This CheapServersCount is not present when CheapServersCount is null.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    null,
                    publicFolderIsGhostedResponse.CheapServersCount,
                    313,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] CheapServersCount (2 bytes): This field is not present if the value of the IsGhosted is zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4631");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4631
                // This Servers is not  present when Servers is null.
                Site.CaptureRequirementIfIsNull(
                    publicFolderIsGhostedResponse.Servers,
                    4631,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] Servers (optional) (variable): This field is not present if IsGhosted is zero.");

                #endregion

                // Step 4: Send a RopDeleteFolder request to the server.
                #region RopDeleteFolder Response
                RopDeleteFolderRequest deleteFolderRequest;

                deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
                deleteFolderRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                deleteFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete;

                // Set FolderId to targetFolderId, this folder is to be deleted.
                deleteFolderRequest.FolderId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopDeleteFolder request.");

                // Send a RopDeleteFolder request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    deleteFolderRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    deleteFolderResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // The ghosted folder is only supported when the second SUT exists.
                if (!string.IsNullOrEmpty(Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site)))
                {
                    // Step 5: Get the existed ghosted public folder ID.
                    folderId = this.GetSubfolderIDByName(openedFolderHandle, Common.GetConfigurationPropertyValue("GhostedPublicFolderDisplayName", this.Site) + "\0");

                    // Step 6: Open the existed ghosted public folder.
                    #region Open folder

                    // Set FolderId to that of existed ghosted public folder.
                    openFolderRequest.FolderId = folderId;

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
                        "if ROP succeeds, the ReturnValue of its response is 0(success)");

                    #endregion

                    #region Verify R560, R564 and R568

                    if (openFolderResponse.IsGhosted > 0x0)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R560");

                        // Verify MS-OXCROPS requirement: MS-OXCROPS_R560
                        // If ServerCount is not null mean present
                        Site.CaptureRequirementIfIsNotNull(
                            openFolderResponse.ServerCount,
                            560,
                            @"[In RopOpenFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is present if IsGhosted is nonzero.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R564");

                        // Verify MS-OXCROPS requirement: MS-OXCROPS_R564
                        // If CheapServerCount is not null mean present
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
                    }

                    #endregion

                    // Step 7: Send RopPublicFolderIsGhosted request to the server and verify RopPublicFolderIsGhosted success response with Ghosted folder.
                    #region RopPublicFolderIsGhosted success response with Ghosted folder

                    // Set FolderId to that of existed ghosted public folder.
                    publicFolderIsGhostedRequest.FolderId = folderId;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the RopPublicFolderIsGhosted request.");

                    this.responseSOHs = cropsAdapter.ProcessSingleRop(
                        publicFolderIsGhostedRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.SuccessResponse);
                    publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)response;

                    Site.Assert.AreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        publicFolderIsGhostedResponse.ReturnValue,
                        "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                    #endregion

                    #region Verify R307,R312 and R316

                    Site.Assert.IsTrue(publicFolderIsGhostedResponse.IsGhosted > 0x0, "If the test case is opening a ghosted public folder, IsGhosted should be greater than 0");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R307");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R307
                    // This ServersCount is present when ServersCount is not null.
                    Site.CaptureRequirementIfIsNotNull(
                        publicFolderIsGhostedResponse.ServersCount,
                        307,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ServersCount (2 bytes): This field[ServersCount (2 bytes)] is present if IsGhosted is nonzero.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R312");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R312
                    // This CheapServersCount is present when CheapServersCount is not null.
                    Site.CaptureRequirementIfIsNotNull(
                        publicFolderIsGhostedResponse.CheapServersCount,
                        312,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] CheapServersCount (2 bytes): This field is present if the value of the IsGhosted field is nonzero.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R316");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R316
                    // This Servers is present when Servers is not null.
                    Site.CaptureRequirementIfIsNotNull(
                        publicFolderIsGhostedResponse.Servers,
                        316,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] Servers (optional) (variable):This field is present if IsGhosted is nonzero.");

                    #endregion

                    if (Common.IsRequirementEnabled(6000101, this.Site))
                    {
                        // Step 8: Create an ghosted public folder of one existed folder.
                        #region Create folder

                        // Set DisplayName to same as that existed folder.
                        createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("GhostedPublicFolderDisplayName", this.Site) + "\0");

                        // Set Comment to same as that existed folder.
                        createFolderRequest.Comment = Encoding.ASCII.GetBytes(Common.GetConfigurationPropertyValue("GhostedPublicFolderComment", this.Site) + "\0");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Step 8: Begin to send the RopCreateFolder request.");

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

                        #region Verify R6000101, R626, R622, R630, R634 and R638

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R6000101");

                        // Verify MS-OXCROPS requirement: MS-OXCROPS_R6000101
                        Site.CaptureRequirementIfAreNotEqual<int>(
                            0,
                            createFolderResponse.IsExistingFolder,
                            6000101,
                            @"[In Appendix A: Product Behavior] If a folder with the name given by the DisplayName field of the request buffer (RopCreateFolder) already exists, implementation does set a nonzero value to IsExistingFolder field. (Exchange 2007 follows this behavior.)");

                        if (createFolderResponse.IsExistingFolder != 0)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R626");

                            // Verify MS-OXCROPS requirement: MS-OXCROPS_R626
                            // IsGhosted not null means present.
                            Site.CaptureRequirementIfIsNotNull(
                                createFolderResponse.IsGhosted,
                                626,
                                @"[In RopCreateFolder ROP Success Response Buffer] IsGhosted (1 byte): This field is present if the value of the IsExistingFolder field is nonzero.");

                            // Refer to MS-OXCSTOR LogonFlags.Private. This bit is set for logon to a private mailbox and is not set for logon to public folders.
                            if (0x00 == (logonResponse.LogonFlags & (byte)LogonFlags.Private))
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R622");

                                // Verify MS-OXCROPS requirement: MS-OXCROPS_R622
                                // HasRules is not null means present.
                                Site.CaptureRequirementIfIsNotNull(
                                    createFolderResponse.HasRules,
                                    622,
                                    @"[In RopCreateFolder ROP Success Response Buffer] HasRules (1 byte): This field is present if the IsExistingFolder field is nonzero.");
                            }

                            if (createFolderResponse.IsGhosted != null && createFolderResponse.IsGhosted != 0)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R630");

                                // Verify MS-OXCROPS requirement: MS-OXCROPS_R630
                                // ServerCount is not null means present.
                                Site.CaptureRequirementIfIsNotNull(
                                    createFolderResponse.ServerCount,
                                    630,
                                    @"[In RopCreateFolder ROP Success Response Buffer] ServerCount (2 bytes): This field is present if the values of both the IsExistingFolder and the IsGhosted fields are nonzero.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R634");

                                // Verify MS-OXCROPS requirement: MS-OXCROPS_R634
                                // CheapServerCount is not null means present.
                                Site.CaptureRequirementIfIsNotNull(
                                    createFolderResponse.CheapServerCount,
                                    634,
                                    @"[In RopCreateFolder ROP Success Response Buffer] CheapServerCount (2 bytes): This field is present if the values of both the IsExistingFolder and the IsGhosted fields are nonzero.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R638");

                                // Verify MS-OXCROPS requirement: MS-OXCROPS_R638
                                // Servers is not null means present.
                                Site.CaptureRequirementIfIsNotNull(
                                    createFolderResponse.Servers,
                                    638,
                                    @"[In RopCreateFolder ROP Success Response Buffer] Servers (variable): This field is present if  the values of both the IsExistingFolder and the IsGhosted fields are nonzero.");
                            }
                        }

                        #endregion
                    }
                }

                // Step 9: Send RopPublicFolderIsGhosted request to the server and verify RopPublicFolderIsGhosted failure response.
                #region RopPublicFolderIsGhosted failure response

                // Set FolderId to 0x1, which does not exist in public folder database and will lead to a failure response.
                publicFolderIsGhostedRequest.FolderId = TestSuiteBase.WrongFolderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 9: Begin to send the RopPublicFolderIsGhosted request.");

                // Send the RopPublicFolderIsGhosted request to the server and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    publicFolderIsGhostedRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    publicFolderIsGhostedResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the first system supports public folder logon.");
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopIdFromLongTermId.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC06_TestRopIdFromLongTermId()
        {
            this.CheckTransportIsSupported();

            // Step 1: Send a RopLongTermIdFromId request to the server and verify the success response.
            #region RopLongTermIdFromId success response

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder and RopCreateFolder request to get the created folder Id.");

            // Create a folder and get its folder id, this id will be converted to a long-term ID.
            ulong objectId = this.GetCreatedsubFolderId(ref logonResponse);

            RopLongTermIdFromIdRequest longTermIdFromIdRequest;
            RopLongTermIdFromIdResponse longTermIdFromIdResponse;

            longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;

            longTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ObjectId to that got in the foregoing code, this id will be converted to a short-term ID.
            longTermIdFromIdRequest.ObjectId = objectId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopLongTermIdFromId request.");

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

            // Step 2: Send a RopIdFromLongTermId to the server and verify the success response.
            #region RopIdFromLongTermId success response

            RopIdFromLongTermIdRequest ropIdFromLongTermIdRequest;
            RopIdFromLongTermIdResponse ropIdFromLongTermIdResponse;

            ropIdFromLongTermIdRequest.RopId = (byte)RopId.RopIdFromLongTermId;
            ropIdFromLongTermIdRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            ropIdFromLongTermIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            ropIdFromLongTermIdRequest.LongTermId.DatabaseGuid = longTermIdFromIdResponse.LongTermId.DatabaseGuid;
            ropIdFromLongTermIdRequest.LongTermId.GlobalCounter = longTermIdFromIdResponse.LongTermId.GlobalCounter;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopIdFromLongTermId request.");

            // Send the RopIdFromLongTermId request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropIdFromLongTermIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            ropIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                ropIdFromLongTermIdResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3: Send a RopIdFromLongTermId to request the server and verify the failure response.
            #region Failure response, verify R471201

            // Set InputHandleIndex to 0x01, which is an invalid index and will lead to a failure response.
            ropIdFromLongTermIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopIdFromLongTermId request.");

            // Send a RopIdFromLongTermId request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropIdFromLongTermIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            ropIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)response;

            if (Common.IsRequirementEnabled(471201, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R471201");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R471201
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.ReturnValueForRopFail,
                    ropIdFromLongTermIdResponse.ReturnValue,
                    471201,
                    @"[In Appendix B: Product Behavior] If the index is invalid, implementation does fail the ROP with the ReturnValue field set to 0x000004B9. (Microsoft Exchange Server 2010 and above follow this behavior.) ");
            }

            #endregion
        }

        /// <summary>
        /// This method tests ROP buffers of RopLongTermIdFromId and RopGetPerUserLongTermIds.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC07_TestRopGetPerUserLongTermIds()
        {
            this.CheckTransportIsSupported();

            // Step 1: Send a RopLongTermIdFromId request to the server and verify the success response.
            #region RopLongTermIdFromId success response

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopLongTermIdFromIdRequest longTermIdFromIdRequest;
            RopLongTermIdFromIdResponse longTermIdFromIdResponse;

            longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;

            longTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ObjectId to the first folder id, this id will be converted to a long-term ID.
            longTermIdFromIdRequest.ObjectId = logonResponse.FolderIds[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopLongTermIdFromId request.");

            // Send a RopLongTermIdFromId request to the server and verify the success response.
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

            // Step 2: Send a RopLongTermIdFromId request to the server and verify the failure response.
            #region RopLongTermIdFromId failure response

            // Set ObjectId to 0x00, which is an invalid id and will lead to a failure response.
            longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopLongTermIdFromId request.");

            // Send a RopLongTermIdFromId request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                longTermIdFromIdRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;

            #endregion

            // Step 3: Send a RopGetPerUserLongTermIds request to the server and verify the success response.
            #region RopGetPerUserLongTermIds success response

            RopGetPerUserLongTermIdsRequest getPerUserLongTermIdsRequest;
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponse;

            getPerUserLongTermIdsRequest.RopId = (byte)RopId.RopGetPerUserLongTermIds;
            getPerUserLongTermIdsRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            getPerUserLongTermIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set DatabaseGuid to that of the LongTermId got in Step 2, this field specifies which database the client is querying data for.
            getPerUserLongTermIdsRequest.DatabaseGuid = longTermIdFromIdResponse.LongTermId.DatabaseGuid;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetPerUserLongTermIds request.");

            // Send a RopGetPerUserLongTermIds request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPerUserLongTermIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPerUserLongTermIdsResponse = (RopGetPerUserLongTermIdsResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPerUserLongTermIdsResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0x00000000 (success)");

            #endregion

            // Check whether the environment supported public folder.
            if (bool.Parse(Common.GetConfigurationPropertyValue("IsPublicFolderSupported", this.Site)))
            {
                // Reconnect to public folder.
                bool ret = this.cropsAdapter.RpcDisconnect();
                this.Site.Assert.IsTrue(ret, "Rpc disconnect should be success.");
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PublicFolderServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));

                // Step 4: Send a RopGetPerUserLongTermIds request to the server and verify the failure response.
                #region RopGetPerUserLongTermIds failure response

                // Log on to a public folder.
                logonResponse = this.Logon(LogonType.PublicFolder, this.userDN, out this.inputObjHandle);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetPerUserLongTermIds request.");

                // Send a RopGetPerUserLongTermIds request to the server and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getPerUserLongTermIdsRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getPerUserLongTermIdsResponse = (RopGetPerUserLongTermIdsResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getPerUserLongTermIdsResponse.ReturnValue,
                    "If ROP fails, the ReturnValue of its response is set to a value other than 0x00000000(failure)");

                #endregion
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopLongTermIdFromId and RopGetPerUserGuid.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC08_TestRopGetPerUserGuid()
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

                // Step 1: Log on to the public folder, and get the public folder ID, we use FolderIds[0]
                RopLogonResponse logonResponse = Logon(LogonType.PublicFolder, this.userDN, out inputObjHandle);
                ulong folderId = logonResponse.FolderIds[2];

                // Step 2: Convert the folderId to longTermID using RopLongTermIdFromId
                #region Convert the folderId to longTermID

                RopLongTermIdFromIdRequest longTermIdFromIdRequest;
                RopLongTermIdFromIdResponse longTermIdFromIdResponse;

                longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
                longTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set ObjectId to the folder id got in Step 1, which will be converted to a long-term ID.
                longTermIdFromIdRequest.ObjectId = folderId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopLongTermIdFromId request.");

                // Send a RopLongTermIdFromId request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    longTermIdFromIdRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)response;
                LongTermId longTermId = longTermIdFromIdResponse.LongTermId;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    longTermIdFromIdResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 3: Log on to a private mailbox.
                this.cropsAdapter.RpcDisconnect();
                this.cropsAdapter.RpcConnect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    ConnectionType.PrivateMailboxServer,
                    Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("PassWord", this.Site));
                logonResponse = this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

                // Step 4: Call RopWritePerUserInformation to set per-user information for the public folder, passing longTermId
                // of the public folder to FolderId filed of RopWritePerUserInformation request.
                #region Set per-user information

                RopWritePerUserInformationRequest writePerUserInformationRequest;
                RopWritePerUserInformationResponse wrtiePerUserInformationResponse;

                // Set per-user information for the public folder by setting the Data field.
                byte[] data = 
                {
                    0xd8, 0x44, 0xae, 0x73, 0xf9, 0x61, 0x5d, 0x4f, 0xb3, 0xc6, 0x9a, 0x7c,
                    0x31, 0xfe, 0xc1, 0x23, 0x06, 0x00, 0x00, 0x00, 0x78, 0x2b, 0x33, 0x00
                };

                writePerUserInformationRequest.RopId = (byte)RopId.RopWritePerUserInformation;
                writePerUserInformationRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                writePerUserInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set FolderId to that got in Step 2, for which folder the per-user information will be set.
                writePerUserInformationRequest.FolderId = longTermId;

                // Set HasFinished to 0xff(TRUE), which indicates this is the last block of data to be returned. 
                writePerUserInformationRequest.HasFinished = TestSuiteBase.NonZero;

                // Set DataOffset to 0, which specifies the location in the per-user information stream to start writing.
                writePerUserInformationRequest.DataOffset = TestSuiteBase.DataOffset;

                // Set DataSize to the length of data, which specifies the size of the Data field in bytes.
                writePerUserInformationRequest.DataSize = (ushort)data.Length;

                // Set Data to data, which will be used to set per-user information for the public folder.
                writePerUserInformationRequest.Data = data;

                // Set ReplicaGuid to that of logonResponse, which identifies which public database is the source of this data. 
                writePerUserInformationRequest.ReplGuid = logonResponse.ReplGuid;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopWritePerUserInformation request.");

                // Send a RopWritePerUserInformation request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    writePerUserInformationRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                wrtiePerUserInformationResponse = (RopWritePerUserInformationResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    wrtiePerUserInformationResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 5:Call RopReadPerUserInformation to verify the data is written successfully.
                #region Verify the data is written successfully

                RopReadPerUserInformationRequest readPerUserInformationRequest;
                RopReadPerUserInformationResponse readPerUserInformationResponse;

                readPerUserInformationRequest.RopId = (byte)RopId.RopReadPerUserInformation;
                readPerUserInformationRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                readPerUserInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set FolderId to the longTermId to that got in Step 2, for which folder the per-user information will be read.
                readPerUserInformationRequest.FolderId = longTermId;

                // Set Reserved, this field is not used and is ignored by the server.
                readPerUserInformationRequest.Reserved = TestSuiteBase.Reserved;

                // Set DataOffset to 0, which specifies the location at which to start reading within the per-user information stream.
                readPerUserInformationRequest.DataOffset = TestSuiteBase.DataOffset;

                // Set MaxDataSize to 30, which specifies the maximum number of bytes of per-user information to be retrieved.
                readPerUserInformationRequest.MaxDataSize = TestSuiteBase.MaxDataSize;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopReadPerUserInformation request.");

                // Send a RopReadPerUserInformation request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    readPerUserInformationRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                readPerUserInformationResponse = (RopReadPerUserInformationResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    readPerUserInformationResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                #endregion

                // Step 6: Call GetPerUserGuid using the longTermID
                RopGetPerUserGuidRequest getPerUserGuidRequest;
                RopGetPerUserGuidResponse getPerUserGuidResponse;

                getPerUserGuidRequest.RopId = (byte)RopId.RopGetPerUserGuid;
                getPerUserGuidRequest.LogonId = TestSuiteBase.LogonId;

                // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                // where the handle for the input Server object is stored.
                getPerUserGuidRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                // Set LongTermId to longTermId, for which folder the GUID of a public folder's per-user information will be got.
                getPerUserGuidRequest.LongTermId = longTermId;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopGetPerUserGuid request.");

                // Send a GetPerUserGuid request to the server and verify the success response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getPerUserGuidRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);
                getPerUserGuidResponse = (RopGetPerUserGuidResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getPerUserGuidResponse.ReturnValue,
                    "If ROP succeeds, the ReturnValue of its response is 0 (success)");

                // Set GlobalCounter and send GetPerUserGuid request again. Here GlobalCounter specifies the folder within its Store object.
                getPerUserGuidRequest.LongTermId.GlobalCounter[0] = Convert.ToByte(TestSuiteBase.GlobalCounter);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 7: Begin to send the GetPerUserGuid request.");

                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getPerUserGuidRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the first system supports public folder logon.");
            }
        }

        /// <summary>
        /// This method tests ROP buffers of RopWritePerUserInformation and RopReadPerUserInformation.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC09_TestRopWriteAndReadPerUserInformation()
        {
            this.CheckTransportIsSupported();

            // Step 1: Send RopLongTermIdFromId request to get the LongTermID, for which folder the per-user information will be set.
            #region Get the LongTermID

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopLongTermIdFromIdRequest longTermIdFromIdRequest;
            RopLongTermIdFromIdResponse longTermIdFromIdResponse;

            longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;

            longTermIdFromIdRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            longTermIdFromIdRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set ObjectId, which will be converted to a long-term ID.
            longTermIdFromIdRequest.ObjectId = logonResponse.FolderIds[4];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopLongTermIdFromId request.");

            // Send a RopLongTermIdFromId request to the server and verify the success response.
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

            // Get longTermId, which will be used in the following RopWritePerUserInformation.
            LongTermId longTermId = longTermIdFromIdResponse.LongTermId;

            #endregion

            // Step 2: Send a RopWritePerUserInformation request to the server and verify the success response.
            #region RopWritePerUserInformation success response

            RopWritePerUserInformationRequest writePerUserInformationRequest;
            RopWritePerUserInformationResponse wrtiePerUserInformationResponse;

            // Set per-user information for the public folder by setting the Data field.
            byte[] data = 
            {
                0xd8, 0x44, 0xae, 0x73, 0xf9, 0x61, 0x5d, 0x4f, 0xb3, 0xc6, 0x9a, 0x7c,
                0x31, 0xfe, 0xc1, 0x23, 0x06, 0x00, 0x00, 0x00, 0x78, 0x2b, 0x33, 0x00
            };
            writePerUserInformationRequest.RopId = (byte)RopId.RopWritePerUserInformation;
            writePerUserInformationRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            writePerUserInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set FolderId to longTermId, for which folder the per-user information will be set.
            writePerUserInformationRequest.FolderId = longTermId;

            // Set HasFinished to 0x0(FALSE), which indicates this is not the last block of data to be returned.
            writePerUserInformationRequest.HasFinished = Convert.ToByte(TestSuiteBase.Zero);

            // Set DataOffset to 0, which specifies the location in the per-user information stream to start writing.
            writePerUserInformationRequest.DataOffset = TestSuiteBase.DataOffset;

            // Set DataSize to the length of the data, which specifies the size of the Data field in bytes.
            writePerUserInformationRequest.DataSize = (ushort)data.Length;

            // Set Data to data, which will be used to set the per-user information of the folder.
            writePerUserInformationRequest.Data = data;

            // Set ReplicaGuid to that of logonResponse, which identifies which public database is the source of this data. 
            writePerUserInformationRequest.ReplGuid = logonResponse.ReplGuid;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopWritePerUserInformation request.");

            // Send a RopWritePerUserInformation request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                writePerUserInformationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            wrtiePerUserInformationResponse = (RopWritePerUserInformationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                wrtiePerUserInformationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3: Send a RopReadPerUserInformation request to the server and verify the success response.
            #region RopReadPerUserInformation success response

            RopReadPerUserInformationRequest readPerUserInformationRequest;
            RopReadPerUserInformationResponse readPerUserInformationResponse;

            readPerUserInformationRequest.RopId = (byte)RopId.RopReadPerUserInformation;
            readPerUserInformationRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            readPerUserInformationRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set FolderId to longTermId, for which folder the per-user information will be read.
            readPerUserInformationRequest.FolderId = longTermId;

            // Set Reserved to 0x00, this field is not used and is ignored by the server.
            readPerUserInformationRequest.Reserved = TestSuiteBase.Reserved;

            // Set DataOffset to 0, which specifies the location at which to start reading within the per-user information stream.
            readPerUserInformationRequest.DataOffset = TestSuiteBase.DataOffset;

            // Set MaxDataSize to 30, which specifies the maximum number of bytes of per-user information to be retrieved.
            readPerUserInformationRequest.MaxDataSize = TestSuiteBase.MaxDataSize;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopReadPerUserInformation request.");

            // Send a RopReadPerUserInformation request to the server and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readPerUserInformationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            readPerUserInformationResponse = (RopReadPerUserInformationResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                readPerUserInformationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            #region Verify R474

            // Change the Reserved field not equal to 0X00 and call this request again.
            readPerUserInformationRequest.Reserved = TestSuiteBase.Reserved;
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readPerUserInformationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopReadPerUserInformationResponse readPerUserInformationResponse1 = (RopReadPerUserInformationResponse)response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R474");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R474
            // Send two different values of Reserved in two RopReadPerUserInformation requests to check whether they have the same return value.
            Site.CaptureRequirementIfAreEqual<uint>(
                readPerUserInformationResponse.ReturnValue,
                readPerUserInformationResponse1.ReturnValue,
                474,
                @"[In RopReadPerUserInformation ROP Request Buffer] Reserved (1 byte): Reply is the same no matter this field is used or not.");

            #endregion

            // Step 4: Send a RopReadPerUserInformation request to the server and verify the failure response.
            #region RopReadPerUserInformation failure response

            // Set DataOffset to more than MaxDataSize to invoke a failure response.
            readPerUserInformationRequest.DataOffset = readPerUserInformationRequest.MaxDataSize++;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopReadPerUserInformation request.");

            // Send a RopReadPerUserInformation request to the server and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                readPerUserInformationRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            readPerUserInformationResponse = (RopReadPerUserInformationResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                readPerUserInformationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests the error for RPC.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC10_TestRPCError()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            RopOpenFolderRequest ropOpenFolderRequest;

            // Step 1: Send a RopOpenFolder request to verify RPC error related to reserved ROP.
            #region The RPC error when it encounters a RopId value that is associated with a reserved ROP

            // Implementation does return an error for the RPC, as specified in [MS-OXCRPC], when it encounters a RopId value 
            // that is associated with a reserved ROP. (Microsoft Exchange Server 2010 and above follow this behavior.)
            if (Common.IsRequirementEnabled(213, this.Site))
            {
                // Reserved RopId array
                byte[] reservedRopIds = 
                {
                    0x00, 0x28, 0x3C, 0x3D, 0x62, 0x65, 0x6A, 0x71, 0x7C, 0x7D, 
                    0x85, 0x87, 0x8A, 0x8B, 0x8C, 0x8D, 0x8E, 0x94, 0x95, 0x96, 
                    0x97, 0x98, 0x99, 0x9A, 0x9B, 0x9C, 0x9D, 0x9E, 0x9F, 0xA0, 
                    0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 
                    0xAB, 0xAC, 0xAD, 0xAE, 0xAF, 0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 
                    0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD, 0xBE, 
                    0xBF, 0xC0, 0xC1, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 
                    0xC9, 0xCA, 0xCB, 0xCC, 0xCD, 0xCE, 0xCF, 0xD0, 0xD1, 0xD2,
                    0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xDB, 0xDC,
                    0xDD, 0xDE, 0xDF, 0xE0, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6,
                    0xE7, 0xE8, 0xE9, 0xEA, 0xEB, 0xEC, 0xED, 0xEE, 0xEF, 0xF0, 
                    0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xFA, 0xFB,
                    0xFC, 0xFD
                };

                foreach (byte ropId in reservedRopIds)
                {
                    // Set a reserved ROP.
                    ropOpenFolderRequest.RopId = ropId;
                    ropOpenFolderRequest.LogonId = TestSuiteBase.LogonId;

                    // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
                    // where the handle for the input Server object is stored.
                    ropOpenFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

                    // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
                    // where the handle for the output Server object will be stored.
                    ropOpenFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

                    // Set FolderId, this folder will be opened.
                    ropOpenFolderRequest.FolderId = logonResponse.FolderIds[4];

                    ropOpenFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopOpenFolder request.");

                    // Send the RopOpenFolder request with a reserved RopId and get the RPC error expected.
                    this.responseSOHs = cropsAdapter.ProcessSingleRop(
                        ropOpenFolderRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.RPCError);
                }
            }

            #endregion

            // Step 2: Send a RopOpenFolder request to verify RPC error related to requests which cannot be parsed.
            #region The RPC error when it encounters the server is unable to parse the ROP requests in the input ROP buffer

            ropOpenFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            ropOpenFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            ropOpenFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x02, which is an invalid index and will lead to the server unable to parse the ROP requests.
            ropOpenFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex2;

            // Set FolderId, this folder will be opened.
            ropOpenFolderRequest.FolderId = logonResponse.FolderIds[4];

            ropOpenFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopOpenFolder request.");

            // Send the RopOpenFolder request and get RPC error.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                ropOpenFolderRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.RPCError);

            #endregion

            if (Common.IsRequirementEnabled(469306, this.Site))
            {
                // Step 3: Send a RopLogon request with lack buffer to verify RPC error ecBufferTooSmall.
                #region Send a RopLogon request with lack buffer to verify RPC error ecBufferTooSmall

                RopLogonRequest logonRequest;

                logonRequest.RopId = (byte)RopId.RopLogon;
                logonRequest.LogonId = TestSuiteBase.LogonId;

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
                this.responseSOHs = cropsAdapter.ProcessSingleRopWithOptionResponseBufferSize(
                            logonRequest,
                            this.inputObjHandle,
                            ref this.response,
                            ref this.rawData,
                            RopResponseType.RPCError,
                            TestSuiteBase.BufferOutOfRange);

                #endregion
            }
        }

        /// <summary>
        /// This method tests the StoreState field in the response and ReturnValues for different StoreState inputs.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S01_TC11_TestStoreStateOfLogon()
        {
            this.CheckTransportIsSupported();

            // Step 1: Set the StoreState to 0 and send the Logon ROP Request.
            #region RopLogon success response

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            RopLogonRequest logonRequest;

            logonRequest.RopId = (byte)RopId.RopLogon;
            logonRequest.LogonId = TestSuiteBase.LogonId;

            // Set OutputHandleIndex to 0x0, which specifies the location in the Server object handle table
            // where the handle for the output Server object will be stored.
            logonRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex0;

            string userDN = Common.GetConfigurationPropertyValue("UserEssdn", this.Site) + "\0";

            // Set StoreState to 0, this field is not used and is ignored by the server.
            logonRequest.StoreState = (uint)StoreState.None;

            logonRequest.LogonFlags = (byte)LogonFlags.Private;
            logonRequest.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;

            // Set EssdnSize to the byte count of user DN, which specifies the size of the Essdn field.
            logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);

            // Set Essdn to the content of user DN, which specifies it will log on to the mail box of user represented by the user DN.
            logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopLogon request.");

            // Send the RopLogon request and get the response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                logonRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopLogonResponse logonResponse = (RopLogonResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                logonResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 2: Set the StoreState to a value other than 0 and send the request again.
            #region Send the request again by setting StoreState to 1

            logonRequest.StoreState = (uint)StoreState.StoreHasSearches;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopLogon request.");

            cropsAdapter.ProcessSingleRop(
                logonRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            RopLogonResponse logonResponse2 = (RopLogonResponse)response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4678");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4678
            Site.CaptureRequirementIfAreEqual<uint>(
                logonResponse.ReturnValue,
                logonResponse2.ReturnValue,
                4678,
                @"[In RopLogon ROP Request Buffer] StoreState (4 bytes): Reply is the same no matter what values are used for this field.");

            #endregion
        }

        #endregion

        #region Common method

        /// <summary>
        /// Get the FolderId of a created subFolder by opening a folder and creating a subfolder.
        /// </summary>
        /// <param name="logonResponse">Logon response</param>
        /// <returns>The created subfolder's ID</returns>
        protected ulong GetCreatedsubFolderId(ref RopLogonResponse logonResponse)
        {
            // Step 1: Open a folder.
            #region Open folder

            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;

            openFolderRequest.LogonId = TestSuiteBase.LogonId;

            // Set InputHandleIndex to 0x0, which specifies the location in the Server object handle table where the handle for the input Server object is stored.
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle for the output Server object will be stored.
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;

            // Set FolderId to 5th folder, which is the FolderId of that will be opened.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            // Get the folder handle, which will be used as input handle in RopCreateFolder.
            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            #endregion

            // Step 2:Create a subfolder under the opened folder
            #region Create subfolder

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

            // Set OpenExisting to 0xFF, which means the folder being created will be opened when it is already existed.
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;

            // Set Reserved to 0x0, this field is reserved and MUST be set to 0.
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
                "If ROP succeeds, the ReturnValue of its response is 0 (success)");

            #endregion

            // Step 3:Get and return the folder id of created folder.
            ulong objectId = createFolderResponse.FolderId;
            return objectId;
        }

        #endregion
    }
}