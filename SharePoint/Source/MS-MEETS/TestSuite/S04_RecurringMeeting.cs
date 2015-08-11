namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 4 Test Cases. Test recurring meeting related requirements.
    /// Add recurring meeting to a workspace.
    /// </summary>
    [TestClass]
    public class S04_RecurringMeeting : TestClassBase
    {
        #region Variables
        /// <summary>
        /// An instance of IMEETSAdapter.
        /// </summary>
        private IMS_MEETSAdapter meetsAdapter;

        /// <summary>
        /// An instance of IMS_MEETSSUTControlAdapter.
        /// </summary>
        private IMS_MEETSSUTControlAdapter sutControlAdapter;
        #endregion

        #region Test suite initialization and cleanup
        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        /// <param name="context">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            // Setup test site
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // Cleanup test site, must be called to ensure closing of logs.
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is used to verify the recurring meeting related requirements.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S04_TC01_RecurringMeetingOperations()
        {
            // Create 3 workspaces.
            string emptyWorkspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(emptyWorkspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");
            string emptyWorkspaceUrl = createWorkspaceResult.Result.CreateWorkspace.Url;

            string singleWorkspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            createWorkspaceResult = this.meetsAdapter.CreateWorkspace(singleWorkspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");
            string singleMeetingWorkspaceUrl = createWorkspaceResult.Result.CreateWorkspace.Url;

            string recurringWorkspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            createWorkspaceResult = this.meetsAdapter.CreateWorkspace(recurringWorkspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");
            string recurringMeetingWorkspaceUrl = createWorkspaceResult.Result.CreateWorkspace.Url;

            // Add a single instance meeting in the single meeting workspace.
            this.meetsAdapter.Url = singleMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeeting = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, meetingTitle, meetingLocation, DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeeting.Exception, "AddMeeting should succeed");

            // Add a recurring meeting in the recurring meeting workspace.
            this.meetsAdapter.Url = recurringMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(Guid.NewGuid().ToString(), true);

            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Send GetMeetingsInformation to the empty workspace.
            this.meetsAdapter.Url = emptyWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> emptyInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(emptyInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.AreEqual<string>("0", emptyInfoResult.Result.MeetingsInformation.WorkspaceStatus.MeetingCount, "Workspace should not contain meeting instance.");

            // Send GetMeetingsInformation to the single meeting workspace.
            this.meetsAdapter.Url = singleMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> singleInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(singleInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.AreEqual<string>("1", singleInfoResult.Result.MeetingsInformation.WorkspaceStatus.MeetingCount, "Workspace should contain only 1 meeting instance.");

            // Send GetMeetingsInformation to the recurring meeting workspace.
            this.meetsAdapter.Url = recurringMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> recurringInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(recurringInfoResult.Exception, "GetMeetingsInformation should succeed");

            // If the MeetingCount that server returns is -1, MS-MEETS_R215 is verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "-1",
                recurringInfoResult.Result.MeetingsInformation.WorkspaceStatus.MeetingCount,
                215,
                @"[In GetMeetingsInformationResponse]This [MeetingCount]MUST be set to -1 if the meeting workspace subsite has a recurring meeting.");         

            // Send GetMeetingWorkspaces to the parent web site.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);

            // Get available workspace for non-recurring meeting, server should return the first and second workspaces.
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspacesResult = this.meetsAdapter.GetMeetingWorkspaces(false);
            Site.Assert.IsNull(getMeetingWorkspacesResult.Exception, "GetMeetingWorkspaces should succeed");

            // If only empty workspaces and single instance workspaces are returned, MS-MEETS_R230 is verified.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length,
                230,
                @"[In GetMeetingWorkspaces][If the value [of recurring]is false], empty workspaces and single instance workspaces are returned.");

            // If only workspaces to which the protocol client can add meetings are returned, MS-MEETS_R224 is verified.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length,
                224,
                @"[In GetMeetingWorkspacesSoapOut]The protocol server MUST return only workspaces to which the protocol client can add meetings.");

            // Get available workspace for recurring meeting, server should only return the first workspace
            getMeetingWorkspacesResult = this.meetsAdapter.GetMeetingWorkspaces(true);
            Site.Assert.IsNull(getMeetingWorkspacesResult.Exception, "GetMeetingWorkspaces should succeed");

            // If only empty workspaces are returned, MS-MEETS_R231 is verified.
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length,
                231,
                @"[In GetMeetingWorkspaces]If [the value of recurring is]true, only empty workspaces are returned.");

            // Clean up the SUT.
            this.meetsAdapter.Url = emptyWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");

            this.meetsAdapter.Url = singleMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");

            this.meetsAdapter.Url = recurringMeetingWorkspaceUrl + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the error code when adding recurring meeting to a non-empty workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S04_TC02_RecurringMeetingError()
        {
            // Create a workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a single instance meeting in the workspace. 
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, meetingTitle, meetingLocation, DateTime.Now, DateTime.Now.AddHours(1), false);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Add a recurring meeting in the same workspace.
            string icalendar = TestSuiteBase.GetICalendar(Guid.NewGuid().ToString(), true);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);

            // Add the log information.
            Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "Verify MS-MEETS_R89: The response when a client tries to add a recurring meeting to a workspace is: {0}", addMeetingFromICalResult.Exception.Detail.InnerText);

            // If a SOAP fault with SOAP fault code "0x00000003" is returned, MS-MEETS_R89 is captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000003",
                addMeetingFromICalResult.GetErrorCode(),
                89,
                @"[In AddMeetingFromICalResponse]If the protocol client tries to add a recurring meeting to a workspace that already contains a meeting, the response MUST be a SOAP fault with SOAP fault code ""0x00000003"".");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }
        #endregion

        #region Test case initialization and cleanup

        /// <summary>
        /// Test case initialize method.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.meetsAdapter = this.Site.GetAdapter<IMS_MEETSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);

            this.sutControlAdapter = Site.GetAdapter<IMS_MEETSSUTControlAdapter>();

            // Make sure the test environment is clean before test case run.
            bool isClean = this.sutControlAdapter.PrepareTestEnvironment(this.meetsAdapter.Url);
            this.Site.Assert.IsTrue(isClean, "The specified site should not have meeting workspaces.");

            // Initialize the TestSuiteBase
            TestSuiteBase.Initialize(this.Site);
        }

        /// <summary>
        /// Test case cleanup method.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.meetsAdapter.Reset();
        }
        #endregion
    }
}