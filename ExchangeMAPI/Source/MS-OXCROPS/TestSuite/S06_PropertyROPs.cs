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
    /// This class is designed to verify the response buffer formats of Property ROPs.
    /// </summary>
    [TestClass]
    public class S06_PropertyROPs : TestSuiteBase
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
        /// This method tests the ROP buffers of RopGetPropertiesAll, RopGetPropertiesList and RopGetPropertiesSpecific.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC01_TestRopsGetProperties()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopGetPropertiesAll request and verify the success response.
            #region RopGetPropertiesAll success response

            // Log on to the private mailbox.
            this.Logon(LogonType.Mailbox, this.userDN, out this.inputObjHandle);

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
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesAllResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 2: Send the RopGetPropertiesAll request and verify the failure response.
            #region RopGetPropertiesAll failure response

            getPropertiesAllRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopGetPropertiesAll request.");

            // Send the RopGetPropertiesAll request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesAllRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesAllResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0");

            #endregion

            // Step 3: Send the RopGetPropertiesList request and verify the success response.
            #region RopGetPropertiesList success response

            RopGetPropertiesListRequest getPropertiesListRequest;
            RopGetPropertiesListResponse getPropertiesListResponse;

            getPropertiesListRequest.RopId = (byte)RopId.RopGetPropertiesList;
            getPropertiesListRequest.LogonId = TestSuiteBase.LogonId;
            getPropertiesListRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopGetPropertiesList request.");

            // Send the RopGetPropertiesList request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesListRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPropertiesListResponse = (RopGetPropertiesListResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesListResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 4: Send the RopGetPropertiesList request and verify the failure response.
            #region RopGetPropertiesList failure response

            getPropertiesListRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetPropertiesList request.");

            // Send the RopGetPropertiesList request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesListRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getPropertiesListResponse = (RopGetPropertiesListResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesListResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0");

            #endregion

            // Step 5: Send the RopGetPropertiesSpecific request and verify the success response.
            #region RopGetPropertiesSpecific success response

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;

            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = TestSuiteBase.LogonId;
            getPropertiesSpecificRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned.
            getPropertiesSpecificRequest.PropertySizeLimit = TestSuiteBase.PropertySizeLimit;

            PropertyTag[] tagArray = this.CreateFolderPropertyTags();
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = tagArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetPropertiesSpecific request.");

            // Send the RopGetPropertiesSpecific request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesSpecificRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesSpecificResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 6: Send the RopGetPropertiesSpecific request and verify the failure response.
            #region RopGetPropertiesSpecific failure response

            getPropertiesSpecificRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopGetPropertiesSpecific request.");

            // Send the RopGetPropertiesSpecific request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertiesSpecificRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertiesSpecificResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0 (success)");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of Test RopCopyTo and RopCopyProperties.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC02_TestRopsCopy()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopCopyTo request and verify the success response.
            #region RopCopyTo success response

            #region Common operations

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetOpenedFolderHandle method to get the opened folder handle.");

            // Call GetOpenedFolderHandle method to get the opened folder handle.
            uint openedFolderHandle = this.GetOpenedFolderHandle(logonResponse.FolderIds[5], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedFolderHandle method to create a folder and get the created folder handle.");

            // Call GetCreatedFolderHandle method to create a folder and get the created folder handle.
            uint createdFolderHandle1 = this.GetCreatedFolderHandle(openedFolderHandle, 1);

            // Call GetCreatedFolderHandle method to create a folder and get the created folder handle.
            uint createdFolderHandle2 = this.GetCreatedFolderHandle(openedFolderHandle, 2);

            List<uint> handleList = new List<uint>
            {
                createdFolderHandle1, createdFolderHandle2
            };

            #endregion

            RopCopyToRequest copyToRequest;
            RopCopyToResponse copyToResponse;

            copyToRequest.RopId = (byte)RopId.RopCopyTo;

            copyToRequest.LogonId = TestSuiteBase.LogonId;

            // Set SourceHandleIndex, which specifies the location in the Server object handle table where the handle
            // for the source Server object is stored.
            copyToRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            // Set DestHandleIndex, which specifies the location in the Server object handle table where the handle
            // for the destination Server object is stored.
            copyToRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;

            copyToRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);
            copyToRequest.WantSubObjects = Convert.ToByte(TestSuiteBase.Zero);
            copyToRequest.CopyFlags = (byte)RopCopyToCopyFlags.NoOverwrite;

            // Set ExcludedTagCount, which specifies how many tags are present in ExcludedTags.
            copyToRequest.ExcludedTagCount = TestSuiteBase.ExcludedTagCount;

            copyToRequest.ExcludedTags = null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopCopyTo request.");

            // Send the RopCopyTo request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyToRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            copyToResponse = (RopCopyToResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyToResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 2: Send the RopCopyTo request and verify the failure response.
            #region RopCopyTo failure response

            // Set SourceHandleIndex and CopyFlags to invalid values, this will lead to a failure response.
            copyToRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex1;
            copyToRequest.CopyFlags = TestSuiteBase.InvalidCopyFlags;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopCopyTo request to invoke the failure response.");

            // Send the RopCopyTo request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyToRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            copyToResponse = (RopCopyToResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyToResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will not be 0");
            Site.Assert.AreNotEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                copyToResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will not be 0x00000503");

            #endregion

            // Step 3: Send the RopCopyTo request and verify the null destination failure response.
            #region RopCopyTo null destination failure response

            // Because the failure response modify its value, so change it to correct.
            copyToRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            handleList.RemoveAt(1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopCopyTo request to invoke null destination failure response.");

            // Send the RopCopyTo request and verify the null destination failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyToRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.NullDestinationFailureResponse);
            copyToResponse = (RopCopyToResponse)response;

            Site.Assert.AreEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                copyToResponse.ReturnValue,
                "if ROP null destination failure response, ReturnValue of its response will be 0x00000503");

            #endregion

            // Step 4: Send the RopCopyProperties request and verify the success response.
            #region RopCopyProperties success response

            RopCopyPropertiesRequest copyPropertiesRequest;
            RopCopyPropertiesResponse copyPropertiesResponse;

            copyPropertiesRequest.RopId = (byte)RopId.RopCopyProperties;
            copyPropertiesRequest.LogonId = TestSuiteBase.LogonId;
            copyPropertiesRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;
            copyPropertiesRequest.DestHandleIndex = TestSuiteBase.DestHandleIndex;
            copyPropertiesRequest.WantAsynchronous = Convert.ToByte(TestSuiteBase.Zero);
            copyPropertiesRequest.CopyFlags = (byte)RopCopyPropertiesCopyFlags.NoOverwrite;

            // Call CreateFolderPropertyTags method to create property tags for folder, then set it to PropertyTags.
            PropertyTag[] tagArray = this.CreateFolderPropertyTags();
            copyPropertiesRequest.PropertyTagCount = (ushort)tagArray.Length;
            copyPropertiesRequest.PropertyTags = tagArray;
            handleList.Add(createdFolderHandle2);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopCopyProperties request.");

            // Send the RopCopyProperties request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyPropertiesRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            copyPropertiesResponse = (RopCopyPropertiesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyPropertiesResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 5: Send the RopCopyProperties request and verify the failure response.
            #region RopCopyProperties failure response

            copyPropertiesRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopCopyProperties request.");

            // Send the RopCopyProperties request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyPropertiesRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            copyPropertiesResponse = (RopCopyPropertiesResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                copyPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0 (success)");
            Site.Assert.AreNotEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                copyPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0x000000503 (null destination failure response)");

            #endregion

            // Step 6: Send the RopCopyProperties request and verify the null destination failure response.
            #region RopCopyProperties null destination failure response

            handleList.RemoveAt(1);

            copyPropertiesRequest.SourceHandleIndex = TestSuiteBase.SourceHandleIndex0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopCopyProperties request.");

            // Send the RopCopyProperties request and verify the null destination failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRopWithMutipleServerObjects(
                copyPropertiesRequest,
                handleList,
                ref this.response,
                ref this.rawData,
                RopResponseType.NullDestinationFailureResponse);
            copyPropertiesResponse = (RopCopyPropertiesResponse)response;

            Site.Assert.AreEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                copyPropertiesResponse.ReturnValue,
                "if ROP null destination failure, ReturnValue of its response will be 0x000000503");

            #endregion
        }

        /// <summary>
        /// This method tests that Exchange 2010 does not support RopProgress.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC03_TestRopProgress()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Refer to MS-OXCPRPT: The initial release version of Exchange 2010 does not implement the RopProgress ROP.
            if (Common.IsRequirementEnabled(86601, this.Site))
            {
                // Step 1: Preparations-Open a folder and construct emptyFolderRequest.
                #region Common methods

                // Log on to a private mailbox.
                RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);
                RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest();
                RopReleaseRequest releaseRequest = new RopReleaseRequest();
                this.PrepareRops(logonResponse, ref createMessageRequest, ref saveChangesMessageRequest, ref releaseRequest);
                uint tableHandle;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1:Call CreateVastMessages method to create Vast Messages In InBox.");

                string transportSeq = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower();
                if (transportSeq == "mapi_http")
                {
                    this.CreateVastMessages(ref logonResponse, out tableHandle, TestSuiteBase.MessagesCount / 50, createMessageRequest, saveChangesMessageRequest, releaseRequest);
                }
                else
                {
                    this.CreateSingleProcessEachLoop(ref logonResponse, out tableHandle, TestSuiteBase.MessagesCount, createMessageRequest, saveChangesMessageRequest, releaseRequest);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 1:Call CreateVastMessages method to create Vast Messages In InBox.");

                // Call GetOpenedFolderHandle to get the opened folder handle.
                uint openedFolderHandle = this.GetOpenedFolderHandle(logonResponse.FolderIds[4], inputObjHandle);

                #endregion

                // Step 2: Verify the RopProgress success response.
                #region RopProgress success response

                // Send the RopEmptyFolder request to delete all messages and subfolders from opened folder.
                #region RopEmptyFolder request

                RopProgressRequest progressRequest;
                RopEmptyFolderRequest emptyFolderRequest;

                emptyFolderRequest.RopId = (byte)RopId.RopEmptyFolder;
                emptyFolderRequest.LogonId = TestSuiteBase.LogonId;
                emptyFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
                emptyFolderRequest.WantAsynchronous = TestSuiteBase.NonZero;
                emptyFolderRequest.WantDeleteAssociated = TestSuiteBase.NonZero;

                #endregion

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopEmptyFolder request.");

                // Send the RopEmptyFolder request.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    emptyFolderRequest,
                    openedFolderHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.SuccessResponse);

                if (response is RopProgressResponse)
                {
                    RopProgressResponse ropProgressResponse = (RopProgressResponse)response;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R86601");

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R86601
                    Site.CaptureRequirementIfAreEqual<uint>(
                        TestSuiteBase.SuccessReturnValue,
                        ropProgressResponse.ReturnValue,
                        86601,
                        @"[In Appendix A: Product Behavior] Implementation does implement the RopProgress ROP ([MS-OXCROPS] section 2.2.8.13). (Exchange 2007 and Exchange 2013 follow this behavior.)");
                }

                #endregion

                if (Common.IsRequirementEnabled(3155, this.Site))
                {
                    // Step 3: Send the RopProgress request and verify the failure response.
                    #region Step 3: Send the RopProgress request and verify the failure response.

                    progressRequest.RopId = (byte)RopId.RopProgress;
                    progressRequest.LogonId = TestSuiteBase.LogonId;
                    progressRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;
                    progressRequest.WantCancel = Convert.ToByte(TestSuiteBase.Zero);

                    Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopProgress request.");

                    // Send the RopProgress request and verify the success response.
                    this.responseSOHs = cropsAdapter.ProcessSingleRop(
                        progressRequest,
                        this.inputObjHandle,
                        ref this.response,
                        ref this.rawData,
                        RopResponseType.FailureResponse);

                    #endregion
                }
            }
        }

        /// <summary>
        /// This method tests the ROP buffers of RopQueryNamedProperties, RopGetPropertyIdsFromNames and RopGetNamesFromPropertyIds.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC04_TestRopsQuery()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Send the RopQueryNamedProperties request and verify the success response.
            #region RopQueryNamedProperties success response

            // Log on to the private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Call GetCreatedMessageHandle to create message and get the created message handle.
            uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            RopQueryNamedPropertiesRequest queryNamedPropertiesRequest = new RopQueryNamedPropertiesRequest();
            RopQueryNamedPropertiesResponse queryNamedPropertiesResponse;

            queryNamedPropertiesRequest.RopId = (byte)RopId.RopQueryNamedProperties;

            queryNamedPropertiesRequest.LogonId = TestSuiteBase.LogonId;
            queryNamedPropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            queryNamedPropertiesRequest.QueryFlags = (byte)QueryFlags.NoStrings;
            queryNamedPropertiesRequest.HasGuid = Convert.ToByte(TestSuiteBase.Zero);
            queryNamedPropertiesRequest.PropertyGuid = null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1: Begin to send the RopQueryNamedProperties request.");

            // Send the RopQueryNamedProperties request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryNamedPropertiesRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            queryNamedPropertiesResponse = (RopQueryNamedPropertiesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryNamedPropertiesResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 2: Send the RopQueryNamedProperties request and verify the failure response.
            #region RopQueryNamedProperties failure response

            queryNamedPropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopQueryNamedProperties request.");

            // Send the RopQueryNamedProperties request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                queryNamedPropertiesRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            queryNamedPropertiesResponse = (RopQueryNamedPropertiesResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                queryNamedPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0 (success)");
            Site.Assert.AreNotEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopQueryNamedProperties,
                queryNamedPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0x00040380 (success)");

            #endregion

            // Step 3: Send the RopGetPropertyIdsFromNames request and verify the success response.
            #region RopGetPropertyIdsFromNames success response

            RopGetPropertyIdsFromNamesRequest getPropertyIdsFromNamesRequest;
            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse;

            getPropertyIdsFromNamesRequest.RopId = (byte)RopId.RopGetPropertyIdsFromNames;
            getPropertyIdsFromNamesRequest.LogonId = TestSuiteBase.LogonId;
            getPropertyIdsFromNamesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            getPropertyIdsFromNamesRequest.Flags = (byte)GetPropertyIdsFromNamesFlags.Create;

            // Call CreatePropertyNameArray method to create propertyName array.
            PropertyName[] propertyNameArray = this.CreatePropertyNameArray(3);

            getPropertyIdsFromNamesRequest.PropertyNameCount = (ushort)propertyNameArray.Length;
            getPropertyIdsFromNamesRequest.PropertyNames = propertyNameArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopQueryNamedProperties request.");

            // Send the RopGetPropertyIdsFromNames request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertyIdsFromNamesRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getPropertyIdsFromNamesResponse = (RopGetPropertyIdsFromNamesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                getPropertyIdsFromNamesResponse.ReturnValue,
                "If RopGetPropertyIdsFromNames succeeds, its response contains ReturnValue 0 (success)");

            #endregion

            // Step 4: Send the RopGetPropertyIdsFromNames request and verify the failure response.
            #region RopGetPropertyIdsFromNames failure response

            getPropertyIdsFromNamesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopGetPropertyIdsFromNames request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getPropertyIdsFromNamesRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion

            // Step 5: Send the RopGetNamesFromPropertyIds request and verify the success response.
            #region RopGetNamesFromPropertyIds success response

            RopGetNamesFromPropertyIdsRequest getNamesFromPropertyIdsRequest;
            RopGetNamesFromPropertyIdsResponse getNamesFromPropertyIdsResponse;

            getNamesFromPropertyIdsRequest.RopId = (byte)RopId.RopGetNamesFromPropertyIds;
            getNamesFromPropertyIdsRequest.LogonId = TestSuiteBase.LogonId;
            getNamesFromPropertyIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            getNamesFromPropertyIdsRequest.PropertyIdCount = getPropertyIdsFromNamesResponse.PropertyIdCount;
            getNamesFromPropertyIdsRequest.PropertyIds = getPropertyIdsFromNamesResponse.PropertyIds;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopGetNamesFromPropertyIds request.");

            // Send the RopGetNamesFromPropertyIds request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                getNamesFromPropertyIdsRequest,
                this.inputObjHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            getNamesFromPropertyIdsResponse = (RopGetNamesFromPropertyIdsResponse)response;

            #endregion

            // Step 6: Send the RopGetNamesFromPropertyIds request and verify the failure response.
            #region RopGetNamesFromPropertyIds failure response

            // Refer to MS-OXCROPS endnote<14>: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the
            // server object and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                getNamesFromPropertyIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;
                PropertyId propertyId = new PropertyId
                {
                    ID = TestSuiteBase.PropertyId
                };
                getNamesFromPropertyIdsRequest.PropertyIdCount = TestSuiteBase.PropertyIdCount;
                getNamesFromPropertyIdsRequest.PropertyIds = new PropertyId[1] { propertyId };

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopGetNamesFromPropertyIds request.");

                // Send the RopGetNamesFromPropertyIds request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getNamesFromPropertyIdsRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getNamesFromPropertyIdsResponse = (RopGetNamesFromPropertyIdsResponse)response;

                Site.Assert.AreEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getNamesFromPropertyIdsResponse.ReturnValue,
                    "<14> Section 3.2.5.1: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.");
            }
            else
            {
                // Verify the response on server other than Exchange 2007.
                getNamesFromPropertyIdsRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;
                PropertyId propertyId = new PropertyId
                {
                    ID = TestSuiteBase.PropertyId
                };
                getNamesFromPropertyIdsRequest.PropertyIdCount = TestSuiteBase.PropertyIdCount;
                getNamesFromPropertyIdsRequest.PropertyIds = new PropertyId[1] { propertyId };

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Step 6: Begin to send the RopGetNamesFromPropertyIds request.");

                // Send the RopGetNamesFromPropertyIds request and verify the failure response.
                this.responseSOHs = cropsAdapter.ProcessSingleRop(
                    getNamesFromPropertyIdsRequest,
                    this.inputObjHandle,
                    ref this.response,
                    ref this.rawData,
                    RopResponseType.FailureResponse);
                getNamesFromPropertyIdsResponse = (RopGetNamesFromPropertyIdsResponse)response;

                Site.Assert.AreNotEqual<uint>(
                    TestSuiteBase.SuccessReturnValue,
                    getNamesFromPropertyIdsResponse.ReturnValue,
                    "For this response, this field SHOULD be set to a value other than 0x00000000.");
            }

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopSetProperties and RopSetPropertiesNoReplicate.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC05_TestRopsSetProperties()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Create message and get the created message handle.
            #region Preparations

            // Log on to a private mailbox.
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedMessageHandle method to create message and get the created message handle.");

            // Call GetCreatedMessageHandle method to create message and get the created message handle.
            uint messageHandle = GetCreatedMessageHandle(logonResponse.FolderIds[4], inputObjHandle);

            int size = 0;
            TaggedPropertyValue[] taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArray(out size);

            #endregion

            // Step 2: Send the RopSetProperties request and verify the success response.
            #region RopSetProperties success response

            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest();
            RopSetPropertiesResponse setPropertiesResponse;

            setPropertiesRequest.RopId = (byte)RopId.RopSetProperties;
            setPropertiesRequest.LogonId = TestSuiteBase.LogonId;
            setPropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            setPropertiesRequest.PropertyValueSize = (ushort)(size + 2);
            setPropertiesRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;
            setPropertiesRequest.PropertyValues = taggedPropertyValueArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopSetProperties request.");

            // Send the RopSetProperties request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setPropertiesRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setPropertiesResponse = (RopSetPropertiesResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setPropertiesResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            #endregion

            // Step 3: Send the RopSetProperties request and verify the failure response.
            #region RopSetProperties failure response

            setPropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopSetProperties request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setPropertiesRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            setPropertiesResponse = (RopSetPropertiesResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0 (success)");
            Site.Assert.AreNotEqual<uint>(
                MS_OXCROPSAdapter.ReturnValueForRopQueryNamedProperties,
                setPropertiesResponse.ReturnValue,
                "if ROP failure, ReturnValue of its response will not be 0x00040380 (success)");

            #endregion

            // Step 4: Send the RopSetPropertiesNoReplicate request and verify the success response.
            #region RopSetPropertiesNoReplicate success response

            RopSetPropertiesNoReplicateRequest setPropertiesNoReplicateRequest = new RopSetPropertiesNoReplicateRequest();
            RopSetPropertiesNoReplicateResponse setPropertiesNoReplicateResponse;

            setPropertiesNoReplicateRequest.RopId = (byte)RopId.RopSetPropertiesNoReplicate;
            setPropertiesNoReplicateRequest.LogonId = TestSuiteBase.LogonId;
            setPropertiesNoReplicateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            setPropertiesNoReplicateRequest.PropertyValueSize = (ushort)(size + 2);
            setPropertiesNoReplicateRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;
            setPropertiesNoReplicateRequest.PropertyValues = taggedPropertyValueArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopSetPropertiesNoReplicate request.");

            // Send the RopSetPropertiesNoReplicate request and verify the success response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setPropertiesNoReplicateRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            setPropertiesNoReplicateResponse = (RopSetPropertiesNoReplicateResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setPropertiesNoReplicateResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");

            #endregion

            // Step 5: Send the RopSetPropertiesNoReplicate request and verify the failure response.
            #region RopSetPropertiesNoReplicate failure response

            setPropertiesNoReplicateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopSetPropertiesNoReplicate request.");

            // Send the RopSetPropertiesNoReplicate request and verify the failure response.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                setPropertiesNoReplicateRequest,
                messageHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);
            setPropertiesNoReplicateResponse = (RopSetPropertiesNoReplicateResponse)response;

            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                setPropertiesNoReplicateResponse.ReturnValue,
                "If ROP failure, the ReturnValue of its response is not 0(success).");

            #endregion
        }

        /// <summary>
        /// This method tests the ROP buffers of RopDeleteProperties and RopDeletePropertiesNoReplicate.
        /// </summary>
        [TestCategory("MSOXCROPS"), TestMethod()]
        public void MSOXCROPS_S06_TC06_TestRopsDeleteProperties()
        {
            this.CheckTransportIsSupported();

            this.cropsAdapter.RpcConnect(
                Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                ConnectionType.PrivateMailboxServer,
                Common.GetConfigurationPropertyValue("UserEssdn", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                Common.GetConfigurationPropertyValue("PassWord", this.Site));

            // Step 1: Preparations-Get the handles of opened folder and created folder.
            #region Preparations

            PropertyTag[] tagArray = this.CreateFolderPropertyTags();
            RopLogonResponse logonResponse = Logon(LogonType.Mailbox, this.userDN, out inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetOpenedFolderHandle method to open a folder and get the opened folder handle.");

            // Call GetOpenedFolderHandle method to open a folder and get the opened folder handle.
            uint openedFolderHandle = this.GetOpenedFolderHandle(logonResponse.FolderIds[4], inputObjHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 1:Call GetCreatedFolderHandle method to create a folder and get the created folder handle.");

            // Call GetCreatedFolderHandle method to create a folder and get the created folder handle.
            uint createdFolderHandle = this.GetCreatedFolderHandle(openedFolderHandle, 1);

            #endregion

            // Step 2: Send the RopDeleteProperties request and verify the success response.
            #region RopDeleteProperties success response

            RopDeletePropertiesRequest deletePropertiesRequest;

            deletePropertiesRequest.RopId = (byte)RopId.RopDeleteProperties;

            deletePropertiesRequest.LogonId = TestSuiteBase.LogonId;
            deletePropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            deletePropertiesRequest.PropertyTagCount = (ushort)tagArray.Length;
            deletePropertiesRequest.PropertyTags = tagArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 2: Begin to send the RopDeleteProperties request.");

            // Send the RopDeleteProperties request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deletePropertiesRequest,
                createdFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion

            // Step 3: Send the RopDeleteProperties request and verify the failure response.
            #region RopDeleteProperties failure response

            deletePropertiesRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 3: Begin to send the RopDeleteProperties request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deletePropertiesRequest,
                createdFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion

            // Step 4: Send the RopDeletePropertiesNoReplicate request and verify the success response.
            #region RopDeletePropertiesNoReplicate success response

            RopDeletePropertiesNoReplicateRequest deletePropertiesNoReplicateRequest;

            deletePropertiesNoReplicateRequest.RopId = (byte)RopId.RopDeletePropertiesNoReplicate;
            deletePropertiesNoReplicateRequest.LogonId = TestSuiteBase.LogonId;
            deletePropertiesNoReplicateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            deletePropertiesNoReplicateRequest.PropertyTagCount = (ushort)tagArray.Length;
            deletePropertiesNoReplicateRequest.PropertyTags = tagArray;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 4: Begin to send the RopDeletePropertiesNoReplicate request.");

            // Send the RopDeletePropertiesNoReplicate request.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deletePropertiesNoReplicateRequest,
                createdFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);

            #endregion

            // Step 5: Send the RopDeletePropertiesNoReplicate request and verify the failure response.
            #region RopDeletePropertiesNoReplicate failure response

            deletePropertiesNoReplicateRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex1;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Step 5: Begin to send the RopDeletePropertiesNoReplicate request.");

            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                deletePropertiesNoReplicateRequest,
                createdFolderHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.FailureResponse);

            #endregion
        }

        #endregion

        #region Common methods

        /// <summary>
        /// Create property tags for folder
        /// </summary>
        /// <returns>Return PropertyTag array</returns>
        private PropertyTag[] CreateFolderPropertyTags()
        {
            PropertyTag[] tags = new PropertyTag[8];

            // PidTagOfflineAddressBookName
            tags[0] = this.propertyDictionary[PropertyNames.PidTagOfflineAddressBookName];

            // PidTagOfflineAddressBookSequence
            tags[1] = this.propertyDictionary[PropertyNames.PidTagOfflineAddressBookSequence];

            // PidTagOfflineAddressBookContainerGuid
            tags[2] = this.propertyDictionary[PropertyNames.PidTagOfflineAddressBookContainerGuid];

            // PidTagOfflineAddressBookMessageClass
            tags[3] = this.propertyDictionary[PropertyNames.PidTagOfflineAddressBookMessageClass];

            // PidTagOfflineAddressDistinguishedName
            tags[4] = this.propertyDictionary[PropertyNames.PidTagOfflineAddressBookDistinguishedName];

            // PidTagSortLocaleId
            tags[5] = this.propertyDictionary[PropertyNames.PidTagSortLocaleId];

            // PidTagMessageCodePage
            tags[6] = this.propertyDictionary[PropertyNames.PidTagMessageCodepage];

            // PidTagEntryId
            tags[7] = this.propertyDictionary[PropertyNames.PidTagEntryId];

            return tags;
        }

        /// <summary>
        /// Create an array of TaggedPropertyValue for message
        /// </summary>
        /// <param name="size">The size of TaggedPropertyValue array</param>
        /// <returns>Return TaggedPropertyValue array</returns>
        private TaggedPropertyValue[] CreateMessageTaggedPropertyValueArray(out int size)
        {
            int arraySize = 0;
            TaggedPropertyValue[] result = new TaggedPropertyValue[2];

            // The following settings are from MS-OXOMSG 4.4.
            // PidTagBody
            result[0] = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagBody].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagBody].PropertyType
                },
                Value = Encoding.Unicode.GetBytes("Sign the order request.\0")
            };

            // PidTagMessageClass
            result[1] = new TaggedPropertyValue
            {
                PropertyTag =
                {
                    PropertyId = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyId,
                    PropertyType = this.propertyDictionary[PropertyNames.PidTagMessageClass].PropertyType
                },
                Value = Encoding.ASCII.GetBytes("50\0")
            };

            for (int i = 0; i < result.Length; i++)
            {
                arraySize += result[i].Size();
            }

            size = arraySize;
            return result;
        }

        /// <summary>
        /// Create propertyName array
        /// </summary>
        /// <param name="propertyNameCount">The count of propertyName</param>
        /// <returns>Return propertyName array</returns>
        private PropertyName[] CreatePropertyNameArray(int propertyNameCount)
        {
            Guid newGUID;
            byte[] unicodeName;
            int lastIndex;

            PropertyName[] propertyNameArray = new PropertyName[propertyNameCount];

            for (int i = 0; i < propertyNameCount; i++)
            {
                newGUID = new Guid();
                propertyNameArray[i] = new PropertyName
                {
                    Kind = 0x01,
                    Guid = newGUID.ToByteArray()
                };

                // The property is identified by the Name field.

                // A Unicode (UTF-16) string, followed by two zero bytes as a null terminator
                // that identifies the property within its property set.
                unicodeName = Encoding.Unicode.GetBytes("ClientDefinedProperty" + i.ToString() + "x");

                lastIndex = unicodeName.Length - 1;
                unicodeName[lastIndex] = 0;
                unicodeName[--lastIndex] = 0;

                // 2-byte null terminator ALSO counted into NameSize.
                propertyNameArray[i].NameSize = (byte)unicodeName.Length;
                propertyNameArray[i].Name = unicodeName;
            }

            return propertyNameArray;
        }

        /// <summary>
        /// Get Opened Folder Handle
        /// </summary>
        /// <param name="folderId">The folder id be used to open folder</param>
        /// <param name="logonHandle">The RopLogon handle</param>
        /// <returns>Return created Message Handle</returns>
        private uint GetOpenedFolderHandle(ulong folderId, uint logonHandle)
        {
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            openFolderRequest.LogonId = TestSuiteBase.LogonId;
            openFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            openFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            openFolderRequest.FolderId = folderId;
            openFolderRequest.OpenModeFlags = (byte)FolderOpenModeFlags.None;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopOpenFolder request in GetOpenedFolderHandle method.");

            // Send the RopOpenFolder request to open a folder.
            this.responseSOHs = cropsAdapter.ProcessSingleRop(
                openFolderRequest,
                logonHandle,
                ref this.response,
                ref this.rawData,
                RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)response;

            Site.Assert.AreEqual<uint>(
                TestSuiteBase.SuccessReturnValue,
                openFolderResponse.ReturnValue,
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            // Get the handle of opened folder.
            uint openedFolderHandle = responseSOHs[0][openFolderResponse.OutputHandleIndex];
            return openedFolderHandle;
        }

        /// <summary>
        /// Get Created Folder Handle.
        /// </summary>
        /// <param name="openedFolderHandle">The opened folder handle</param>
        /// <param name="tempFolderIndex">The temp folder's index</param>
        /// <returns>Return created Folder Handle</returns>
        private uint GetCreatedFolderHandle(uint openedFolderHandle, int tempFolderIndex)
        {
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest();
            RopCreateFolderResponse createFolderResponse;

            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = TestSuiteBase.LogonId;
            createFolderRequest.InputHandleIndex = TestSuiteBase.InputHandleIndex0;
            createFolderRequest.OutputHandleIndex = TestSuiteBase.OutputHandleIndex1;
            createFolderRequest.FolderType = (byte)FolderType.Genericfolder;
            createFolderRequest.UseUnicodeStrings = Convert.ToByte(TestSuiteBase.Zero);
            createFolderRequest.OpenExisting = TestSuiteBase.NonZero;
            createFolderRequest.Reserved = TestSuiteBase.Reserved;

            // Set DisplayName, which specifies the name of the created folder.
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + tempFolderIndex + "\0");

            // Set Comment, which specifies the folder comment that is associated with the created folder.
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(TestSuiteBase.DisplayNameAndCommentForNonSearchFolder + tempFolderIndex + "\0");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Begin to send the RopCreateFolder request in GetCreatedFolderHandle method.");

            // Send the RopCreateFolder request to create folder.
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
                "if ROP succeeds, ReturnValue of its response will be 0 (success)");

            // Get the handle of the created folder.
            uint createdFolderHandle = responseSOHs[0][createFolderResponse.OutputHandleIndex];
            return createdFolderHandle;
        }

        #endregion
    }
}