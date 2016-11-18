namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 3 Test Cases. Test iCalendar meeting related operations and requirements.
    /// Include adding and updating meeting to a workspace based on a calendar object.
    /// </summary>
    [TestClass]
    public class S03_MeetingFromICal : TestClassBase
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
        /// This test case is used to test typical meeting based on calendar object scenario.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC01_MeetingFromICalOperations()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            string icalendar = TestSuiteBase.GetICalendar(uid, false, meetingTitle, meetingLocation);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update the meeting.
            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalRecurring = TestSuiteBase.GetICalendar(uid, true, meetingTitle, meetingLocation, attendeeEmail);
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(icalRecurring, false);
            Site.Assert.IsNull(updateMeetingFromICalResult.Exception, "UpdateMeetingFromICal should succeed");

            // If the returned status code equals to 0, MS-MEETS_R32 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                0,
                updateMeetingFromICalResult.Result.UpdateMeetingFromICal.AttendeeUpdateStatus.Code,
                32,
                @"[In AttendeeUpdateStatus]This number [Code]is set to zero, if there was no error.");

            // If UpdateMeetingFromICal executed succeed and there is no error, the returned Detail value is empty, MS-MEETS_R34 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                updateMeetingFromICalResult.Result.UpdateMeetingFromICal.AttendeeUpdateStatus.Detail,
                34,
                @"[In AttendeeUpdateStatus]This string [Detail]is empty, if there was no error.");

            // If the returned status code equals to 0 and the returned Detail value is empty, MS-MEETS_R41110 can be verified.
            Site.CaptureRequirement(
                41110,
                @"[In UpdateMeetingFromICal]If icalText is present, the UpdateMeetingFromICal operation will succeed.");

            // update attendee response.
            SoapResult<Null> setAttendeeResponsResult = this.meetsAdapter.SetAttendeeResponse(attendeeEmail, 0, uid, 0, DateTime.Now.AddHours(1), DateTime.Now.AddHours(2), AttendeeResponse.responseAccepted);
            Site.Assert.IsNull(setAttendeeResponsResult.Exception, "SetAttendeeResponse should succeed");

            // Clean up the SUT. 
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the error when the operations AddMeetingFromICal, UpdateMeetingFromICal and SetAttendeeResponse are sent to a web site that is not a meeting workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC02_MeetingFromICalInvalidUrlError()
        {
            // Set the Url to the default site, which is not workspace.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(Guid.NewGuid().ToString(), false);

            // Send AddMeetingFromICal to a web site that is not a workspace.
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNotNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal operation failed.");

            // If the returned SOAP fault contains the error code "0x00000006", MS-MEETS_R88 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000006",
                addMeetingFromICalResult.GetErrorCode(),
                88,
                @"[In AddMeetingFromICalResponse]If this operation [AddMeetingFromICal] is sent to a Web site that is not a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000006"".");

            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(icalendar, null);
            Site.Assert.IsNotNull(updateMeetingFromICalResult.Exception, "UpdateMeetingFromICal should fail.");

            // If the returned SOAP fault contains the error code "0x00000006", MS-MEETS_R380 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000006",
                updateMeetingFromICalResult.GetErrorCode(),
                380,
                @"[In UpdateMeetingFromICalResponse]If this operation [UpdateMeetingFromICal] is sent to a web site that is not a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000006"".");

            // Send SetAttendeeResponseSoapIn to a website that is not a workspace.
            SoapResult<Null> setAttendeeResponseResult = this.meetsAdapter.SetAttendeeResponse(organizerEmail, null, Guid.NewGuid().ToString(), null, null, null, null);
            Site.Assert.IsNotNull(setAttendeeResponseResult.Exception, "SetAttendeeResponse operation failed.");

            // If server returns a SOAP fault, MS-MEETS_R297 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                setAttendeeResponseResult.Exception,
                297,
                @"[In SetAttendeeResponseResponse]The SetAttendeeResponseResponse element contains nothing other than standard SOAP faults if an error occurs.");
        }

        /// <summary>
        /// This test case is used to verify the error of UpdateMeeting with invalid parameter.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC03_UpdateMeetingFromICalError()
        {
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add a meeting to the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, TestSuiteBase.GetICalendar(Guid.NewGuid().ToString(), false));
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Send UpdateMeetingFromICal to the workspace with icalText parameter set to empty.
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> emptyICalResult = this.meetsAdapter.UpdateMeetingFromICal(string.Empty, null);

            // If the returned SOAP fault contains error code "0x00000005", MS-MEETS_R367, MS-MEETS_R378 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000005",
                emptyICalResult.GetErrorCode(),
                367,
                @"[In UpdateMeetingFromICal]If this parameter [icalText] is an empty string, the response MUST be a SOAP fault with SOAP fault code ""0x00000005"".");

            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000005",
                emptyICalResult.GetErrorCode(),
                378,
                @"[In UpdateMeetingFromICalResponse]If the icalText value is empty, the protocol server returns a SOAP fault with SOAP fault code ""0x00000005"".");

            // Since UpdateMeetingFromICal operation failed in above steps and with SOAP fault code "0x00000005" returned, so MS-MEETS_R377 can be captured directly.
            Site.CaptureRequirement(
                377,
                @"[In UpdateMeetingFromICalResponse]If the operation [UpdateMeetingFromICal]is not successful, this [UpdateMeetingFromICalResponse]represents a SOAP fault.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test AddMeetingFromICal operation when the parameter icalText contains more than 254 attendees' elements.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC04_AddMeetingFromICalWithInvalidAttendees()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar when the parameter icalText contains more than 254 attendees elements.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string icalendar = TestSuiteBase.GetICalendar(uid, false, TestSuiteBase.GenerateAttendees(255));
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNotNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // If icalText contains more than 254 attendees and server returned error code "0x0000000d", MS-MEETS_R4004 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x0000000d",
                addMeetingFromICalResult.GetErrorCode(),
                4004,
                @"[In AddMeetingFromICal]If this parameter[icalText] contains more than 254 attendee elements, the response MUST be a SOAP fault with SOAP fault code ""0x0000000d"".");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test SetAttendeeResponse operation when server ignores utcDateTimeOrganizerCriticalChange element and uses only utcDateTimeAttendeeCriticalChange.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC05_SetAttendeeResponseForIgnoreUtcDateTimeOrganizerCriticalChange()
        {
            // Verify the production is Windows Share Point 3.0.
            if (Common.IsRequirementEnabled(4100, this.Site))
            {
                string uid = Guid.NewGuid().ToString();

                // Create a workspace on the server.
                string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
                SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
                Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

                // Add meeting from ICalendar.
                this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

                string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
                string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
                string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
                SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
                Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

                // Update attendee response with different utcDateTimeOrganizerCriticalChange element.
                SoapResult<Null> setAttendeeResponsResult = this.meetsAdapter.SetAttendeeResponse(attendeeEmail, 0, uid, 0, DateTime.Now.AddHours(1), DateTime.Now.AddHours(2), AttendeeResponse.responseAccepted);
                Site.Assert.IsNull(setAttendeeResponsResult.Exception, "SetAttendeeResponse should succeed");

                SoapResult<Null> setAttendeeResponsResultAgain = this.meetsAdapter.SetAttendeeResponse(attendeeEmail, 0, uid, 0, DateTime.Now.AddHours(2), DateTime.Now.AddHours(2), AttendeeResponse.responseAccepted);
                Site.Assert.IsNull(setAttendeeResponsResult.Exception, "SetAttendeeResponse should succeed");

                // If input two different utcDateTimeOrganizerCriticalChange values, the returned responses are the same, MS-MEETS_R4100 can be verified.
                Site.CaptureRequirementIfAreEqual<Null>(
                   setAttendeeResponsResult.Result,
                   setAttendeeResponsResultAgain.Result,
                   4100,
                   @"[In Appendix B: Product Behavior]For two different utcDateTimeOrganizerCriticalChange values, the implementation of the protocol server response will be same.(Windows SharePoint Service 3.0 follows this behavior)");

                // Clean up the SUT. 
                this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
                SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
                Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
            }
        }

        /// <summary>
        /// This test case is used to test UpdateMeetingFromICal operation when the parameter icalText contains more than 254 attendees' elements.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC06_UpdateMeetingFromICalWithInvalidAttendees()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update meeting from ICalendar when the parameter icalText contains more than 254 attendees elements.
            string icalendarUpdate = TestSuiteBase.GetICalendar(uid, false, TestSuiteBase.GenerateAttendees(255));
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(icalendarUpdate, false);
            Site.Assert.IsNotNull(updateMeetingFromICalResult.Exception, "UpdateMeetingFromICal should fail.");

            // If the returned SOAP fault contains the error code "0x0000000d", MS-MEETS_R4009 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x0000000d",
                updateMeetingFromICalResult.GetErrorCode(),
                4009,
                @"[In UpdateMeetingFromICal]If this parameter [icalText]contains more than 254 ATTENDEE tags, the response MUST be a SOAP fault with SOAP fault code ""0x0000000d"".");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test UpdateMeetingFromICal operation when the parameter icalText is empty.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC07_UpdateMeetingFromICalWithEmptyicalText()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(4114, this.Site), "This case runs only when the requirement 4114 is enabled.");
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string icalendar = TestSuiteBase.GetICalendar(uid, false);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update meeting from ICalendar when the parameter icalText is empty.
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(null, false);

            // Set the "icalTest" to empty and execute the UpdateMeetingFromICal operation, if the returned exception is not null, it means the server returns a SOAP fault, so MS-MEETS_R4114 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                updateMeetingFromICalResult.Exception,
                4114,
                @"[In Appendix B: Product Behavior]Implementation does return a SOAP fault. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test add meeting from icalendar with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC08_AddMeetingFromICalWithAllParametersSpecified()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar with all parameters specified.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test add meeting from icalendar without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC09_AddMeetingFromICalWithoutOptionalParameters()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar without optional parameters.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(null, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test SetAttendeeResponse operation with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC10_SetAttendeeResponseWithAllParametersSpecified()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Set attendee response with all parameters specified.
            SoapResult<Null> setAttendeeResponsResult = this.meetsAdapter.SetAttendeeResponse(attendeeEmail, 0, uid, 0, DateTime.Now.AddHours(2), DateTime.Now.AddHours(2), AttendeeResponse.responseAccepted);
            Site.Assert.IsNull(setAttendeeResponsResult.Exception, "SetAttendeeResponse should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test SetAttendeeResponse operation without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC11_SetAttendeeResponseWithoutOptionalParameters()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Set attendee response without optional parameters.
            SoapResult<Null> setAttendeeResponsResult = this.meetsAdapter.SetAttendeeResponse(attendeeEmail, null, uid, null, null, null, null);
            Site.Assert.IsNull(setAttendeeResponsResult.Exception, "SetAttendeeResponse should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test UpdateMeetingFromICal operation with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC12_UpdateMeetingFromICalWithAllParametersSpecified()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update the meeting with all parameters specified.
            string updateICal = TestSuiteBase.GetICalendar(uid, false, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), attendeeEmail);
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(updateICal, false);
            Site.Assert.IsNull(updateMeetingFromICalResult.Exception, "UpdateMeetingFromICal should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test UpdateMeetingFromICal operation without optional parameters.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC13_UpdateMeetingFromICalWithoutOptionalParameters()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update the meeting without optional parameters.
            string updateICal = TestSuiteBase.GetICalendar(uid, false, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), attendeeEmail);
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(updateICal, null);
            Site.Assert.IsNull(updateMeetingFromICalResult.Exception, "UpdateMeetingFromICal should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test UpdateMeetingFromICal operation when icalText is not present.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC14_UpdateMeetingFromICalWhenICalTextNotPresent()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string attendeeEmail = Common.GetConfigurationPropertyValue("AttendeeEmail", this.Site);
            string icalendar = TestSuiteBase.GetICalendar(uid, false, attendeeEmail);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            Site.Assert.IsNull(addMeetingFromICalResult.Exception, "AddMeetingFromICal should succeed");

            // Update the meeting when icalText is not present.
            SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> updateMeetingFromICalResult = this.meetsAdapter.UpdateMeetingFromICal(null, false);

            // If server returns an exception, that is to say, when the icalText is not present, the UpdateMeetingFromICal operation will failed with an exception. Then MS-MEETS_R41111 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                updateMeetingFromICalResult.Exception,
                41111,
                @"[In UpdateMeetingFromICal]If icalText is not present, the UpdateMeetingFromICal operation will failed with an exception.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test SOAP fault "0x0000000a" is returned if calls AddMeetingFromICal with an invalid organizerEmail.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S03_TC15_MeetingFromICalWithInvalidOrganizerEmail()
        {
            string uid = Guid.NewGuid().ToString();

            // Create a workspace on the server
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add meeting from ICalendar.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            string icalendar = TestSuiteBase.GetICalendar(uid, false, meetingTitle, meetingLocation);
            string organizerEmail = Common.GenerateResourceName(this.Site, "InvalidOrganizerEmail");
            SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> addMeetingFromICalResult = this.meetsAdapter.AddMeetingFromICal(organizerEmail, icalendar);
            string errorCode = Common.ExtractErrorCodeFromSoapFault(addMeetingFromICalResult.Exception);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-MEETS_R29971");

            // Verify MS-VERSS requirement: MS-MEETS_R29971
            Site.CaptureRequirementIfAreEqual<string>(
                "0x0000000a",
                errorCode,
                29971,
                @"[In AddMeetingFromICal]If this parameter [organizerEmail] is an invalid e-mail address, the response MUST be a SOAP fault with SOAP fault code ""0x0000000a"".");

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

            string url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            this.meetsAdapter.Url = url;

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