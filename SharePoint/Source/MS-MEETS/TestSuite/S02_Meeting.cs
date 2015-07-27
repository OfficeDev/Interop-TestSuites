//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 2 Test Cases. Test meeting related operations and requirements.
    /// Include adding meeting, updating meeting, deleting and restoring the meeting.
    /// </summary>
    [TestClass]
    public class S02_Meeting : TestClassBase
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
        /// This test case is used to test the typical meeting scenario.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC01_MeetingOperations()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, uid, null, DateTime.Now, meetingTitle, meetingLocation, DateTime.Now, DateTime.Now.AddHours(1), false);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // If the Url returned by AddMeeting is a well formatted Uri string, MS-MEETS_R24 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString("addMeetingResult.Result.AddMeeting.Url", UriKind.RelativeOrAbsolute),
                24,
                @"[In AddMeeting]Url: The absolute URL of the meeting instance in the workspace, with an indicator of the instance in the absolute URL  query section.");

            string updatedMeetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string updatedLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<Null> updateMeetingResult = this.meetsAdapter.UpdateMeeting(uid, 1, null, updatedMeetingTitle, updatedLocation, DateTime.Now.AddHours(1), DateTime.Now.AddHours(2), null);
            Site.Assert.IsNull(updateMeetingResult.Exception, "UpdateMeeting should succeed");
          
            // Remove the meeting.
            SoapResult<Null> removeMeetingResult = this.meetsAdapter.RemoveMeeting(null, uid, 1, null, null);
            Site.Assert.IsNull(removeMeetingResult.Exception, "RemoveMeeting should succeed");

            // Get workspace status to query the meeting information.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getMeetingsInformationResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getMeetingsInformationResult.Exception, "GetMeetingsInformation should succeed");

            // Restore the removed meeting.
            SoapResult<Null> restoreMeetingResult = this.meetsAdapter.RestoreMeeting(uid);
            Site.Assert.IsNull(restoreMeetingResult.Exception, "RestoreMeeting should succeed");

            // Get workspace status again to query the meeting information.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getMeetingsInformationAgainResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getMeetingsInformationAgainResult.Exception, "GetMeetingsInformation should succeed");

            // Clean up the SUT
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the error when restoring a non-existent meeting.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC02_RestoreMeetingError()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Restore a meeting which does not exist.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> restoreMeetingResult = this.meetsAdapter.RestoreMeeting(uid);

            // If Restore a un-existed meeting failed with returned "0x8102003e" error code, MS-MEETS_R2937 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x8102003e",
                restoreMeetingResult.GetErrorCode(),
                2937,
                @"[In RestoreMeetingResponse]If the meeting specified by the uid parameter in the RestoreMeeting operation does not exist in the meeting workspace, a SOAP fault response is returned with SOAP fault code 0x8102003e.");

            // Clean up the SUT
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case verifies that server returns a SOAP fault when the AddMeeting and UpdateMeeting operations were sent to a web site that is not a meeting workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC03_MeetingInvalidUrlError()
        {
            string uid = Guid.NewGuid().ToString();

            // Set the Url to the default site, which is not workspace.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, uid, null, DateTime.Now, meetingTitle, meetingLocation, DateTime.Now, DateTime.Now.AddHours(1), false);

            // If error code "0x00000006" is returned, MS-MEETS_R66 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000006",
                addMeetingResult.GetErrorCode(),
                66,
                @"[In AddMeetingResponse]If this operation [AddMeeting]is sent to a Web site (2) that is not a meeting workspace, the response [AddMeetingResponse]MUST be a SOAP fault with SOAP Fault code ""0x00000006"".");

            string updatedMeetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string updatedLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<Null> updateMeetingResult = this.meetsAdapter.UpdateMeeting(uid, null, null, updatedMeetingTitle, updatedLocation, DateTime.Now, DateTime.Now.AddHours(1), null);

            // If error code "0x00000006" is returned, MS-MEETS_R350 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000006",
                updateMeetingResult.GetErrorCode(),
                350,
                @"[In UpdateMeetingResponse]If this operation [UpdateMeeting] is sent to a Web site (2) that is not a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000006"".");
        }

        /// <summary>
        /// This test case is used to verify the meeting count under workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC04_VerifyMeetingCountInWorkspace()
        {
            // Create a workspace on site.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitleFst = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocationFst = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResultFst = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, meetingTitleFst, meetingLocationFst, DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResultFst.Exception, "Add meeting should succeed");

            // Get workspace status, make sure there is only one meeting in workspace.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResultFst = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getWorkspaceInfoResultFst.Exception, "Get meeting information should succeed");
            Site.Assert.AreEqual<string>("1", getWorkspaceInfoResultFst.Result.MeetingsInformation.WorkspaceStatus.MeetingCount, "There is only one meeting in workspace");

            // Add another meeting in workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string meetingTitleSnd = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocationSnd = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResultSnd = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, meetingTitleSnd, meetingLocationSnd, DateTime.Now, DateTime.Now.AddHours(2), null);
            Site.Assert.IsNull(addMeetingResultSnd.Exception, "Add meeting should succeed");

            // Get workspace status, make sure there are two meetings in workspace.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResultSnd = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getWorkspaceInfoResultSnd.Exception, "Get meeting information should succeed");
            string actualMeetingCount = getWorkspaceInfoResultSnd.Result.MeetingsInformation.WorkspaceStatus.MeetingCount;

            // If the returned meeting count equals to 2, MS-MEETS_R27 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                actualMeetingCount,
                27,
                @"[In AddMeeting]MeetingCount: The number of meeting instances in the workspace, including the one just added.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test add meeting with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC05_AddMeetingWithAllParametersSpecified()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting with all parameters specified. 
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), 1, DateTime.Now, meetingTitle, meetingLocation, DateTime.Now, DateTime.Now.AddHours(1), false);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test add meeting without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC06_AddMeetingWithoutOptionalParameters()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting without optional parameters.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(null, Guid.NewGuid().ToString(), null, null, null, null, DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test update meeting with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC07_UpdateMeetingWithAllParametersSpecified()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(null, uid, null, null, null, null, DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Update the meeting with all parameters specified.
            string updatedMeetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string updatedLocation = TestSuiteBase.GetUniqueMeetingLocation();
            SoapResult<Null> updateMeetingResult = this.meetsAdapter.UpdateMeeting(uid, 1, DateTime.Now, updatedMeetingTitle, updatedLocation, DateTime.Now.AddHours(1), DateTime.Now.AddHours(2), false);
            Site.Assert.IsNull(updateMeetingResult.Exception, "UpdateMeeting should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test update meeting without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC08_UpdateMeetingWithoutOptionalParameters()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(null, uid, null, null, null, null, DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Update the meeting without optional parameters.
            SoapResult<Null> updateMeetingResult = this.meetsAdapter.UpdateMeeting(uid, null, null, null, null, DateTime.Now.AddHours(1), DateTime.Now.AddHours(2), null);
            Site.Assert.IsNull(updateMeetingResult.Exception, "UpdateMeeting should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test remove meeting with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC09_RemoveMeetingWithAllParametersSpecified()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site), uid, null, null, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Remove the meeting with all parameters specified.
            SoapResult<Null> removeMeetingResult = this.meetsAdapter.RemoveMeeting(0, uid, 1, DateTime.Now, true);
            Site.Assert.IsNull(removeMeetingResult.Exception, "RemoveMeeting should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test remove meeting without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S02_TC10_RemoveMeetingWithoutOptionalParameters()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site), uid, null, null, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            // Remove the meeting with all parameters specified.
            SoapResult<Null> removeMeetingResult = this.meetsAdapter.RemoveMeeting(null, uid, null, null, null);
            Site.Assert.IsNull(removeMeetingResult.Exception, "RemoveMeeting should succeed");

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