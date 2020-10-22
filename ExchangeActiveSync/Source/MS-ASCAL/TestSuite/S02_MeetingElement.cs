namespace Microsoft.Protocols.TestSuites.MS_ASCAL
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// This scenario is to test Calendar class elements, which are attached in a Meeting request, when meeting is either accepted, tentative accepted, cancelled or declined.
    /// </summary>
    [TestClass]
    public class S02_MeetingElement : TestSuiteBase
    {
        #region Test Class initialize and clean up

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        #region MSASCAL_S02_TC01_MeetingAccepted

        /// <summary>
        /// This case is designed to verify ResponseType, AttendeeStatus, Name, Email, AppointmentReplyTime , BusyStatus, MeetingStatus and AttendeeType when recipient accepts the meeting.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC01_MeetingAccepted()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.BusyStatus, (byte)0
                },
                {
                    Request.ItemsChoiceType8.MeetingStatus, (byte)1
                },
                {
                    Request.ItemsChoiceType8.Attendees,
                    TestSuiteHelper.CreateAttendeesRequired(
                        new string[]
                        {
                            Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain)
                        },
                        new string[]
                        {
                            this.User2Information.UserName
                        })
                }
            };

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            // Sync command to do an initialization Sync, and get the organizer calendars changes before attendee accepting the meeting
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R11311");

            // Verify MS-ASCAL requirement: MS-ASCAL_R11311
            Site.CaptureRequirementIfIsTrue(
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified,
                11311,
                @"[In AttendeeStatus] The AttendeeStatus element specifies the attendee's acceptance status.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R568");

            // Verify MS-ASCAL requirement: MS-ASCAL_R568
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatus,
                568,
                @"[In AttendeeStatus] [The value is] 0 [meaning] Response unknown.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R11911");

            // Verify MS-ASCAL requirement: MS-ASCAL_R11911
            Site.CaptureRequirementIfIsTrue(
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeTypeSpecified,
                11911,
                @"[In AttendeeType] The AttendeeType element specifies whether the attendee is required, optional, or a resource.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R573");

            // Verify MS-ASCAL requirement: MS-ASCAL_R573
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeType,
                573,
                @"[In AttendeeType] [The value is] 1 [meaning] Required.");

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.BusyStatus, "The BusyStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13811");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13811
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                calendarOnOrganizer.Calendar.BusyStatus.Value,
                13811,
                @"[In BusyStatus] [The value is] 0 [meaning] Free.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13311");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13311
            // MS-ASCAL_R13811 can be captured. Along with the logic, MS-ASCAL_R13311 can be captured also.
            Site.CaptureRequirement(
                13311,
                @"[In BusyStatus] The BusyStatus element specifies the busy status of the meeting organizer.");

            if (!this.IsActiveSyncProtocolVersion121)
            {
                Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.ResponseType, "The ResponseType element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R40111");

                // Verify MS-ASCAL requirement: MS-ASCAL_R40111
                // If Calendar.ResponseType is not null, it means the server returns the type of response made by the user to a meeting request
                Site.CaptureRequirementIfIsNotNull(
                    calendarOnOrganizer.Calendar.ResponseType.Value,
                    40111,
                    @"[In ResponseType] As a top-level element of the Calendar class, the ResponseType<17> element specifies the type of response made by the user to a meeting request.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R410");

                // Verify MS-ASCAL requirement: MS-ASCAL_R410
                Site.CaptureRequirementIfAreEqual<uint>(
                    1,
                    calendarOnOrganizer.Calendar.ResponseType.Value,
                    410,
                    @"[In ResponseType] [The value 1 means]Organizer. The current user is the organizer of the meeting and, therefore, no reply is required.");
            }

            #endregion

            #region Switch to attendee to accept the meeting request, and sync calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes before accepting the meeting
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            if (calendarOnAttendee.Calendar != null)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            }
            else
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);
            }

            // Accept the meeting request and send the respond to organizer
            bool isSuccess = this.MeetingResponse(byte.Parse("1"), this.User2Information.InboxCollectionId, emailItem.ServerId, null);

            if (isSuccess)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, subject);
            }
            else
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.Fail("MeetingResponse command failed.");
            }

            this.SendMimeMeeting(calendarOnAttendee.Calendar, subject, Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), "REPLY", "ACCEPTED");

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes after accepting the meeting
            calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar.ResponseType, "The ResponseType element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R414");

                // Verify MS-ASCAL requirement: MS-ASCAL_R414
                Site.CaptureRequirementIfAreEqual<uint>(
                    3,
                    calendarOnAttendee.Calendar.ResponseType.Value,
                    414,
                    @"[In ResponseType] [The value 3 means] Accepted. The user has accepted the meeting request.");

                Site.Assert.IsNotNull(calendarOnAttendee.Calendar.AppointmentReplyTime, "The value of AppointmentReplyTime element should not be null since the attendee has replied this meeting request.");

                // Update Location value
                SyncStore syncResponse = this.SyncChanges(this.User2Information.CalendarCollectionId);
                Dictionary<Request.ItemsChoiceType7, object> changeItem = new Dictionary<Request.ItemsChoiceType7, object>();
                string newLocation = Common.GenerateResourceName(Site, "newLocation");
                changeItem.Add(Request.ItemsChoiceType7.Location1, newLocation);
                changeItem.Add(Request.ItemsChoiceType7.Subject, subject);

                this.UpdateCalendarProperty(calendarOnAttendee.ServerId, this.User2Information.CalendarCollectionId, syncResponse.SyncKey, changeItem);

                SyncItem newCalendar = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

                Site.Assert.IsNotNull(newCalendar.Calendar, "The updated calendar should be found.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R96");

                // Verify MS-ASCAL requirement: MS-ASCAL_R96
                Site.CaptureRequirementIfAreEqual<string>(
                    calendarOnAttendee.Calendar.AppointmentReplyTime.Value.ToString("yyyyMMddTHHmmssZ"),
                    newCalendar.Calendar.AppointmentReplyTime.Value.ToString("yyyyMMddTHHmmssZ"),
                    96,
                    @"[In AppointmentReplyTime] The top-level AppointmentReplyTime element can be ghosted.");
            }

            #endregion

            #region Switch to organizer to call Sync command to sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            // Call Sync command to do an initialization Sync, and get the organizer inbox changes
            emailItem = this.GetChangeItem(this.User1Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The organizer should have received the response.");

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, subject);

            // Sync command to do an initialization Sync, and get the organizer calendars changes after attendee accepted the meeting
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R570");

            // Verify MS-ASCAL requirement: MS-ASCAL_R570
            Site.CaptureRequirementIfIsTrue(
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified && calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatus == 3,
                570,
                @"[In AttendeeStatus] [The value is] 3 [meaning]Accept.");
        }

        #endregion

        #region MSASCAL_S02_TC02_MeetingDeclined

        /// <summary>
        /// This test case is designed to verify ResponseType, AttendeeStatus, Name, Email, AppointmentReplyTime , BusyStatus, MeetingStatus and AttendeeType when recipient declines the meeting.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC02_MeetingDeclined()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.BusyStatus, (byte)2
                },
                {
                    Request.ItemsChoiceType8.MeetingStatus, (byte)1
                }
            };

            Request.Attendees attendees = TestSuiteHelper.CreateAttendeesRequired(new string[] { Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain) }, new string[] { this.User2Information.UserName });
            attendees.Attendee[0].AttendeeType = 2;
            attendees.Attendee[0].AttendeeTypeSpecified = true;
            calendarItem.Add(Request.ItemsChoiceType8.Attendees, attendees);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            // Sync command to do an initialization Sync, and get the organizer calendars changes before attendee declining the meeting
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);
            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.BusyStatus, "The BusyStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13813");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13813
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                calendarOnOrganizer.Calendar.BusyStatus.Value,
                13813,
                @"[In BusyStatus] [The value is] 2 [meaning] Busy.");

            if (!this.IsActiveSyncProtocolVersion121)
            {
                Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.ResponseRequested, "The ResponseRequested element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R39611");

                // Verify MS-ASCAL requirement: MS-ASCAL_R39611
                Site.CaptureRequirementIfIsNotNull(
                    calendarOnOrganizer.Calendar.ResponseRequested.Value,
                    39611,
                    @"[In ResponseRequested] The ResponseRequested<16> element specifies whether a response to the meeting request is required.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R574");

            // Verify MS-ASCAL requirement: MS-ASCAL_R574
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeType,
                574,
                @"[In AttendeeType] [The value is] 2 [meaning] Optional.");

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R311");

            // Verify MS-ASCAL requirement: MS-ASCAL_R311
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarOnOrganizer.Calendar.MeetingStatus.Value,
                311,
                @"[In MeetingStatus][The value 1 means] The event is a meeting and the user is the meeting organizer.");

            #endregion

            #region Switch to attendee to decline the meeting request, and sync calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes before declining the meeting
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            if (calendarOnAttendee.Calendar == null)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);
            }

            // Decline the meeting request and send the response to organizer
            bool isSuccess = this.MeetingResponse(byte.Parse("3"), this.User2Information.InboxCollectionId, emailItem.ServerId, null);

            if (!isSuccess)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.Fail("MeetingResponse command failed.");
            }

            this.SendMimeMeeting(calendarOnAttendee.Calendar, subject, Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), "REPLY", "DECLINED");

            // Use EmptyFolderContents to empty the Deleted Items folder
            this.DeleteAllItems(this.User2Information.DeletedItemsCollectionId);

            #endregion

            #region Switch to organizer to call Sync command to Sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            // Call Sync command to do an initialization Sync, and get the organizer inbox changes
            emailItem = this.GetChangeItem(this.User1Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The organizer should have received the response.");

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, subject);

            // Sync command to do an initialization Sync, and get the organizer calendars changes
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R571");

            // Verify MS-ASCAL requirement: MS-ASCAL_R571
            Site.CaptureRequirementIfIsTrue(
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified = true && calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatus == 4,
                571,
                @"[In AttendeeStatus] [The value is] 4 [meaning] Decline.");
        }

        #endregion

        #region MSASCAL_S02_TC03_MeetingTentative

        /// <summary>
        /// This test case is designed to verify ResponseType, AttendeeStatus, Name, Email, AppointmentReplyTime , BusyStatus, MeetingStatus and AttendeeType when the meeting is tentative.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC03_MeetingTentative()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.BusyStatus, (byte)1
                },
                {
                    Request.ItemsChoiceType8.MeetingStatus, (byte)1
                }
            };

            Request.Attendees attendees = TestSuiteHelper.CreateAttendeesRequired(new string[] { Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain) }, new string[] { this.User2Information.UserName });

            calendarItem.Add(Request.ItemsChoiceType8.Attendees, attendees);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            // Sync command to do an initialization Sync, and get the organizer calendars changes before attendee tentative the meeting
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);
            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.BusyStatus, "The BusyStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13812");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13812
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarOnOrganizer.Calendar.BusyStatus.Value,
                13812,
                @"[In BusyStatus] [The value is] 1 [meaning] Tentative.");

            #endregion

            #region Switch to attendee to tentative the meeting request, and sync calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes before the meeting is tentative
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            if (calendarOnAttendee.Calendar != null)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            }
            else
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);
            }

            // Tentatively accept the meeting request and send the response to organizer
            bool isSuccess = this.MeetingResponse(byte.Parse("2"), this.User2Information.InboxCollectionId, emailItem.ServerId, null);

            if (isSuccess)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, subject);
            }
            else
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.Fail("MeetingResponse command failed.");
            }

            this.SendMimeMeeting(calendarOnAttendee.Calendar, subject, Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), "REPLY", "TENTATIVE");

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes after the meeting is tentative
            calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);
            Site.Assert.IsNotNull(calendarOnAttendee.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R312");

            // Verify MS-ASCAL requirement: MS-ASCAL_R312
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                calendarOnAttendee.Calendar.MeetingStatus.Value,
                312,
                @"[In MeetingStatus][The value 3 means] This event is a meeting and the user is not the meeting organizer; the meeting was received from someone else.");

            if (!this.IsActiveSyncProtocolVersion121)
            {
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar.ResponseType, "The ResponseType element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R412");

                // Verify MS-ASCAL requirement: MS-ASCAL_R412
                Site.CaptureRequirementIfAreEqual<uint>(
                    2,
                    calendarOnAttendee.Calendar.ResponseType.Value,
                    412,
                    @"[In ResponseType] [The value 2 means] Tentative. The user is unsure whether he or she will attend.");
            }

            #endregion

            #region Switch to organizer to call Sync command to sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            // Call Sync command to do an initialization Sync, and get the organizer inbox changes
            emailItem = this.GetChangeItem(this.User1Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The organizer should have received the response.");

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, subject);

            // Sync command to do an initialization Sync, and get the organizer calendars changes
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R569");

            // Verify MS-ASCAL requirement: MS-ASCAL_R569
            Site.CaptureRequirementIfIsTrue(
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatusSpecified && calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatus == 2,
                569,
                @"[In AttendeeStatus] [The value is] 2 [meaning]Tentative.");
        }

        #endregion

        #region MSASCAL_S02_TC04_MeetingNotResponded

        /// <summary>
        /// This test case is designed to verify ResponseType, ResponseRequested, AttendeeStatus, Name, Email, AppointmentReplyTime, BusyStatus, MeetingStatus and AttendeeType when recipient respond the meeting with no action.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC04_MeetingNotResponded()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.BusyStatus, (byte)3
                },
                {
                    Request.ItemsChoiceType8.MeetingStatus, (byte)1
                }
            };

            Request.Attendees attendees = TestSuiteHelper.CreateAttendeesRequired(new string[] { Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain) }, new string[] { this.User2Information.UserName });

            attendees.Attendee[0].AttendeeType = 3;
            attendees.Attendee[0].AttendeeTypeSpecified = true;
            if (this.IsActiveSyncProtocolVersion121
                || this.IsActiveSyncProtocolVersion140
                || this.IsActiveSyncProtocolVersion141)
            {
                attendees.Attendee[0].AttendeeStatusSpecified = true;
                attendees.Attendee[0].AttendeeStatus = 5;
            }

            calendarItem.Add(Request.ItemsChoiceType8.Attendees, attendees);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            // Sync command to do an initialization Sync, and get the organizer calendars changes after the meeting request sent
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R575");

            // Verify MS-ASCAL requirement: MS-ASCAL_R575
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeType,
                575,
                @"[In AttendeeType] [The value is] 3 [meaning] Resource.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13814");

            // Verify MS-ASCAL requirement: MS-ASCAL_R13814
            Site.CaptureRequirementIfAreEqual<byte>(
                3,
                calendarOnOrganizer.Calendar.BusyStatus.Value,
                13814,
                @"[In BusyStatus] [The value is] 3 [meaning] OutofOffice.");

            #endregion

            #region Switch to attendee to Sync calendars from the server

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);

            // Respond the meeting with no action, only call Sync command to do an initialization Sync, and get the attendee calendars changes
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar.ResponseType, "The ResponseType element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R418");

                // Verify MS-ASCAL requirement: MS-ASCAL_R418
                Site.CaptureRequirementIfAreEqual<uint>(
                    5,
                    calendarOnAttendee.Calendar.ResponseType.Value,
                    418,
                    @"[In ResponseType] [The value 5 means] Not Responded. The user has not yet responded to the meeting request.");
            }

            #endregion

            #region Switch to organizer to call Sync command to sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            // Sync command to do an initialization Sync, and get the organizer calendars changes
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            if (!this.IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52522");

                // Verify MS-ASCAL requirement: MS-ASCAL_R52522
                // If AppointmentReplyTime is null, it means the server does not return the date and time that the current user responded to the meeting request
                Site.CaptureRequirementIfIsNull(
                    calendarOnOrganizer.Calendar.AppointmentReplyTime,
                    52522,
                    @"[In Message Processing Events and Sequencing Rules][The following information pertains to all command responses:] If no action has been taken on a meeting request, the server MUST NOT include the AppointmentReplyTime element as a top-level element in a command response.");
            }

            if (this.IsActiveSyncProtocolVersion121
                || this.IsActiveSyncProtocolVersion140
                || this.IsActiveSyncProtocolVersion141)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R572");

                // Verify MS-ASCAL requirement: MS-ASCAL_R572
                Site.CaptureRequirementIfAreEqual<byte>(
                    5,
                    calendarOnOrganizer.Calendar.Attendees.Attendee[0].AttendeeStatus,
                    572,
                    @"[In AttendeeStatus] [The value is] 5 [meaning] Not responded.");
            }
        }

        #endregion

        #region MSASCAL_S02_TC05_MeetingCancellation

        /// <summary>
        /// This test case is designed to verify element MeetingStatus when meeting is cancelled.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC05_MeetingCancellation()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>
            {
                {
                    Request.ItemsChoiceType8.MeetingStatus, (byte)5
                },
                {
                    Request.ItemsChoiceType8.Attendees,
                    TestSuiteHelper.CreateAttendeesRequired(
                        new string[]
                        {
                            Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain)
                        },
                        new string[]
                        {
                            this.User2Information.UserName
                        })
                }
            };

            if (!this.IsActiveSyncProtocolVersion121)
            {
                calendarItem.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            #endregion

            #region Switch to attendee to sync mail and calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            // Respond the meeting with no action, only call Sync command to do an initialization Sync, and get the attendee calendars changes
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            if (calendarOnAttendee.Calendar == null)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);
            }

            #endregion

            #region Switch to organizer to send a cancel meeting request.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            this.SendMimeMeeting(calendarOnOrganizer.Calendar, "CANCELLED:" + subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "CANCEL", null);

            #endregion

            #region Switch to attendee to sync calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            emailItem = this.GetChangeItem(this.User2Information.InboxCollectionId, "CANCELLED:" + subject);
            Site.Assert.AreEqual<string>(
                "CANCELLED:" + subject,
                emailItem.Email.Subject,
                "The attendee should have received the cancel response.");

            // Sync command to do an initialization Sync, and get the attendee calendars changes after organizer cancelled meeting
            calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, "CANCELLED:" + subject);

            if (calendarOnAttendee.Calendar != null)
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, "CANCELLED:" + subject);
            }
            else
            {
                this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, "CANCELLED:" + subject);
                Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", "CANCELLED:" + subject);
            }

            // Use EmptyFolderContents to empty the attendee's Inbox folder
            this.DeleteAllItems(this.User2Information.InboxCollectionId);

            #endregion

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R314");

            // Verify MS-ASCAL requirement: MS-ASCAL_R314
            Site.CaptureRequirementIfAreEqual<byte>(
                7,
                calendarOnAttendee.Calendar.MeetingStatus.Value,
                314,
                @"[In MeetingStatus][The value 7 means] The meeting has been canceled. The user was not the meeting organizer, the meeting was received from someone else.");

            #region Switch to organizer to call Sync command to sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, true);

            // Sync command to do an initialization Sync, and get the organizer calendars changes
            calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            if (this.IsActiveSyncProtocolVersion121
                || this.IsActiveSyncProtocolVersion140
                || this.IsActiveSyncProtocolVersion141
                || this.IsActiveSyncProtocolVersion160
                || this.IsActiveSyncProtocolVersion161)
            {
                Site.Assert.IsNotNull(calendarOnOrganizer.Calendar.MeetingStatus, "The MeetingStatus element should not be null.");
                
                
                
                // Add the debug information 
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R313");

                // Verify MS-ASCAL requirement: MS-ASCAL_R313
                Site.CaptureRequirementIfAreEqual<byte>(
                    5,
                    calendarOnOrganizer.Calendar.MeetingStatus.Value,
                    313,
                    @"[In MeetingStatus][The value 5 means] The meeting has been canceled and the user was the meeting organizer.");
                
            }
        }

        #endregion

        #region MSASCAL_S02_TC06_ExceptionElements

        /// <summary>
        /// This test case is designed to verify all elements in Exception.
        /// </summary>
        [TestCategory("MSASCAL"), TestMethod()]
        public void MSASCAL_S02_TC06_ExceptionElements()
        {
            #region Organizer calls Sync command to add a calendar to the server, and sync calendars from the server.

            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            Dictionary<Request.ItemsChoiceType8, object> calendarItem = new Dictionary<Request.ItemsChoiceType8, object>();

            DateTime exceptionStartTime = this.StartTime.AddDays(2);
            DateTime startTimeInException = exceptionStartTime.AddMinutes(15);
            DateTime endTimeInException = startTimeInException.AddHours(2);

            // Set Calendar StartTime, EndTime elements
            calendarItem.Add(Request.ItemsChoiceType8.StartTime, this.StartTime.ToString("yyyyMMddTHHmmssZ"));
            calendarItem.Add(Request.ItemsChoiceType8.EndTime, this.EndTime.ToString("yyyyMMddTHHmmssZ"));

            // Set Calendar BusyStatus element
            calendarItem.Add(Request.ItemsChoiceType8.BusyStatus, (byte)1);

            // Set Calendar Attendees element with required sub-element
            calendarItem.Add(Request.ItemsChoiceType8.Attendees, TestSuiteHelper.CreateAttendeesRequired(new string[] { Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain) }, new string[] { this.User2Information.UserName }));

            // Set Calendar Recurrence element including Occurrence sub-element
            byte recurrenceType = byte.Parse("0");
            calendarItem.Add(Request.ItemsChoiceType8.Recurrence, this.CreateCalendarRecurrence(recurrenceType, 6, 1));

            // Set Calendar Exceptions element
            Request.Exceptions exceptions = new Request.Exceptions { Exception = new Request.ExceptionsException[] { } };
            List<Request.ExceptionsException> exceptionList = new List<Request.ExceptionsException>();

            // Set ExceptionStartTime element in exception
            Request.ExceptionsException exception1 = TestSuiteHelper.CreateExceptionRequired(exceptionStartTime.ToString("yyyyMMddTHHmmssZ"));

            exception1.StartTime = startTimeInException.ToString("yyyyMMddTHHmmssZ");
            exception1.EndTime = endTimeInException.ToString("yyyyMMddTHHmmssZ");
            exception1.Attendees = TestSuiteHelper.CreateAttendeesRequired(new string[] { Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "tester@test.com" }, new string[] { this.User2Information.UserName, "test" }).Attendee;

            exception1.Subject = "Calendar Exception";
            exception1.Body = TestSuiteHelper.CreateCalendarBody(2, this.Content + "InException");
            exception1.BusyStatusSpecified = true;
            exception1.BusyStatus = 2;
            exception1.Location = "Room 666";
            exception1.Reminder = 10;
            exceptionList.Add(exception1);
            exceptions.Exception = exceptionList.ToArray();
            calendarItem.Add(Request.ItemsChoiceType8.Exceptions, exceptions);

            string subject = Common.GenerateResourceName(Site, "subject");
            calendarItem.Add(Request.ItemsChoiceType8.Subject, subject);

            this.AddSyncCalendar(calendarItem);

            SyncItem calendar = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendar.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);

            #endregion

            #region Organizer sends the meeting request to attendee.

            this.SendMimeMeeting(calendar.Calendar, subject, Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain), Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain), "REQUEST", null);

            // Sync command to do an initialization Sync, and get the organizer calendars changes after the meeting request sent
            SyncItem calendarOnOrganizer = this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnOrganizer.Calendar, "The calendar with subject {0} should exist in server.", subject);

            #endregion

            if (!this.IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R10611");

                // Verify MS-ASCAL requirement: MS-ASCAL_R10611
                Site.CaptureRequirementIfIsTrue(
                    calendarOnOrganizer.Calendar.Exceptions.Exception[0].Attendees != null,
                    10611,
                    @"[In Attendees] The Attendees element specifies the collection of attendees for the calendar item exception.<2>");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R39211");

            // Verify MS-ASCAL requirement: MS-ASCAL_R39211
            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(calendarOnOrganizer.Calendar.Exceptions.Exception[0].Reminder.ToString()) == false,
                39211,
                @"[In Reminder] The Reminder element specifies the number of minutes before a calendar item exception's start time to display a reminder notice.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R12611");

            // Verify MS-ASCAL requirement: MS-ASCAL_R12611
            // If Exception.Body is not null, it means the server returns the body text of the calendar item exception
            Site.CaptureRequirementIfIsNotNull(
                calendarOnOrganizer.Calendar.Exceptions.Exception[0].Body,
                12611,
                @"[In Body (AirSyncBase Namespace)] The airsyncbase:Body element specifies the body text of the calendar item exception.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R13411");
        
            // Verify MS-ASCAL requirement: MS-ASCAL_R13411
            this.Site.CaptureRequirementIfAreEqual<byte?>(
                calendar.Calendar.BusyStatus,
                calendarOnOrganizer.Calendar.BusyStatus,
                13411,
                @"[In BusyStatus] A command response has a maximum of one BusyStatus child element per Exception element.");

            #region Switch to attendee to accept the meeting request, and sync calendars from the server.

            // Switch to attendee
            this.SwitchUser(this.User2Information, true);

            // Call Sync command to do an initialization Sync, and get the attendee inbox changes
            SyncItem emailItem = GetChangeItem(this.User2Information.InboxCollectionId, subject);
            Site.Assert.AreEqual<string>(
                subject,
                emailItem.Email.Subject,
                "The attendee should have received the meeting request.");

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes before accepting the meeting
            SyncItem calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);

            // Respond the meeting request
            #region Accept the fourth occurrence

            this.MeetingResponse(byte.Parse("1"), this.User2Information.CalendarCollectionId, calendarOnAttendee.ServerId, startTimeInException.ToString("yyyy-MM-ddThh:mm:ss.000Z"));

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes after accepting the meeting
            calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);

            if (!this.IsActiveSyncProtocolVersion121)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R89011");

                // Verify MS-ASCAL requirement: MS-ASCAL_R89011
                // If Exception.AppointmentReplyTime is not null, it means the server returns the date and time that the user responded to the meeting request exception
                Site.CaptureRequirementIfIsNotNull(
                    calendarOnAttendee.Calendar.Exceptions.Exception[0].AppointmentReplyTime,
                    89011,
                    @"[In AppointmentReplyTime] The AppointmentReplyTime element specifies the date and time that the user responded to the meeting request exception.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R40211");

                // Verify MS-ASCAL requirement: MS-ASCAL_R40211
                Site.CaptureRequirementIfIsTrue(
                    calendarOnAttendee.Calendar.Exceptions.Exception[0].ResponseTypeSpecified,
                    40211,
                    @"[In ResponseType] The ResponseType<18> element specifies the type of response made by the user to a recurring meeting exception.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R23611");

            // Verify MS-ASCAL requirement: MS-ASCAL_R23611
            // If Exception.EndTime is not null, it means the server returns the end time of the calendar item exception
            Site.CaptureRequirementIfIsNotNull(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].EndTime,
                23611,
                @"[In EndTime] The EndTime element specifies the end time of the calendar item exception.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R29611");

            // Verify MS-ASCAL requirement: MS-ASCAL_R29611
            // If Exception.Location is not null, it means the server returns the place where the event specified by the calendar item exception occurs
            Site.CaptureRequirementIfIsNotNull(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].Location,
                29611,
                @"[In Location] The Location element specifies the place where the event specified by the calendar item exception occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R43511");

            // Verify MS-ASCAL requirement: MS-ASCAL_R43511
            // If Exception.StartTime is not null, it means the server returns the start time of the calendar item exception
            Site.CaptureRequirementIfIsNotNull(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].StartTime,
                43511,
                @"[In StartTime] The StartTime element specifies the start time of the calendar item exception.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R44111");

            // Verify MS-ASCAL requirement: MS-ASCAL_R44111
            // If Exception.Subject is not null, it means the server returns the subject of the calendar item exception
            Site.CaptureRequirementIfAreEqual<string>(
                "Calendar Exception".ToLower(CultureInfo.CurrentCulture),
                calendarOnAttendee.Calendar.Exceptions.Exception[0].Subject.ToLower(CultureInfo.CurrentCulture),
                44111,
                @"[In Subject] The Subject element specifies the subject of the calendar item exception.");

            if (!this.IsActiveSyncProtocolVersion121 && !this.IsActiveSyncProtocolVersion140)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R30511");

                // Verify MS-ASCAL requirement: MS-ASCAL_R30511
                Site.CaptureRequirementIfIsTrue(
                    calendarOnAttendee.Calendar.Exceptions.Exception[0].MeetingStatusSpecified,
                    30511,
                    @"[In MeetingStatus] The MeetingStatus element specifies the status of the calendar item exception.");
            }

            #endregion

            #region Decline the fifth occurrence

            this.MeetingResponse(byte.Parse("3"), this.User2Information.CalendarCollectionId, calendarOnAttendee.ServerId, startTimeInException.AddDays(1).ToString("yyyy-MM-ddThh:mm:ss.000Z"));

            // Call Sync command to do an initialization Sync, and get the attendee calendars changes after accepting the meeting
            calendarOnAttendee = this.GetChangeItem(this.User2Information.CalendarCollectionId, subject);

            Site.Assert.IsNotNull(calendarOnAttendee.Calendar, "The calendar with subject {0} should exist in server.", subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R26111");

            // Verify MS-ASCAL requirement: MS-ASCAL_R26111
            // If Exceptions element is not null, it means the server returns a collection of exceptions to the recurrence pattern of the calendar item
            Site.CaptureRequirementIfIsNotNull(
                calendarOnAttendee.Calendar.Exceptions,
                26111,
                @"[In Exceptions] The Exceptions element specifies a collection of exceptions to the recurrence pattern of the calendar item.");

            foreach (Response.ExceptionsException exception in calendarOnAttendee.Calendar.Exceptions.Exception)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R24111");

                // Verify MS-ASCAL requirement: MS-ASCAL_R24111
                // If Exceptions.Exception is not null, it means the server returns an exception to the calendar item's recurrence pattern
                Site.CaptureRequirementIfIsNotNull(
                    exception,
                    24111,
                    @"[In Exception] The Exception element specifies an exception to the calendar item's recurrence pattern.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R20611");

            // Verify MS-ASCAL requirement: MS-ASCAL_R20611
            Site.CaptureRequirementIfIsTrue(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].DeletedSpecified,
                20611,
                @"[In Deleted] The Deleted element specifies whether the exception to the calendar item has been deleted.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R20811");

            // Verify MS-ASCAL requirement: MS-ASCAL_R20811
            Site.CaptureRequirementIfIsTrue(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].DeletedSpecified,
                20811,
                @"[In Deleted] A command response has a maximum of one Deleted child element per Exception element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R210");

            // Verify MS-ASCAL requirement: MS-ASCAL_R210
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                calendarOnAttendee.Calendar.Exceptions.Exception[0].Deleted,
                210,
                @"[In Deleted] An exception will be deleted when the Deleted element is included as a child element of the Exception element with a value of 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R525221");

            // Verify MS-ASCAL requirement: MS-ASCAL_R525221
            // If Exception.AppointmentReplyTime is null, it means the server does not return the date and time that the user responded to the meeting request exception
            Site.CaptureRequirementIfIsNull(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].AppointmentReplyTime,
                525221,
                @"[In Message Processing Events and Sequencing Rules][The following information pertains to all command responses:] If a meeting request exception has not been accepted, the server MUST NOT include the AppointmentReplyTime element as a child element of the Exception element in a command response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCAL_R52522111");

            // Verify MS-ASCAL requirement: MS-ASCAL_R52522111
            // If Exception.AppointmentReplyTime is null, it means the server does not return the date and time that the user responded to the meeting request exception
            Site.CaptureRequirementIfIsNull(
                calendarOnAttendee.Calendar.Exceptions.Exception[0].AppointmentReplyTime,
                52522111,
                @"[In Message Processing Events and Sequencing Rules][The following information pertains to all command responses:] If a meeting request exception has not been tentatively accepted, the server MUST NOT include the AppointmentReplyTime element as a child element of the Exception element in a command response.");

            #endregion

            #endregion

            #region Switch to organizer to call Sync command to sync calendars from the server.

            // Switch to organizer
            this.SwitchUser(this.User1Information, false);

            // Sync command to do an initialization Sync, and get the organizer calendars changes
            this.GetChangeItem(this.User1Information.CalendarCollectionId, subject);

            #endregion
        }

        #endregion

        #endregion
    }
}