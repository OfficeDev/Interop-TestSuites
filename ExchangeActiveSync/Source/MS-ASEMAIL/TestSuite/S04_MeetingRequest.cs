namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test meeting request events, including sending a meeting request to server, synchronizing the meeting request with server.
    /// </summary>
    [TestClass]
    public class S04_MeetingRequest : TestSuiteBase
    {
        #region Class initialize and clean up
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
        public static void ClassCleanUp()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region MSASEMAIL_S04_TC01_SyncAdd_CalendarItem
        /// <summary>
        /// This case is designed to test when a user creates an appointment or meeting on the client, the calendar item will be added to the server by using the Sync command.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC01_SyncAdd_CalendarItem()
        {
            #region Call Sync command with Add element to add an appointment to the server
            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData();

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType8> itemsElementName = new List<Request.ItemsChoiceType8>();

            string subject = Common.GenerateResourceName(Site, "Subject");
            items.Add(subject);
            itemsElementName.Add(Request.ItemsChoiceType8.Subject);

            // MeetingStauts is set to 0, which means it is an appointment with no attendees.
            byte meetingStatus = 0;
            items.Add(meetingStatus);
            itemsElementName.Add(Request.ItemsChoiceType8.MeetingStatus);

            applicationData.Items = items.ToArray();
            applicationData.ItemsElementName = itemsElementName.ToArray();

            SyncStore initialSync = this.InitializeSync(this.User1Information.CalendarCollectionId);
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(initialSync.SyncKey, this.User1Information.CalendarCollectionId, applicationData);

            SyncStore syncAddResponse = this.EMAILAdapter.Sync(syncAddRequest);
            Site.Assert.IsTrue(syncAddResponse.AddResponses[0].Status.Equals("1"), "The sync add operation should be success; It is:{0} actually", syncAddResponse.AddResponses[0].Status);

            // Add the appointment to clean up list.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the new added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R14");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R14
            Site.CaptureRequirementIfAreEqual<string>(
                subject,
                item.Calendar.Subject,
                14,
                @"[In Sending and Receiving Meeting Requests] When a user creates an appointment on the client, the calendar item is added to the server by using the Sync command ([MS-ASCMD] section 2.2.1.21).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC02_MeetingRequest_NoRecurrence
        /// <summary>
        /// This case is designed to test the meeting request which has no recurrence.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC02_MeetingRequest_NoRecurrence()
        {
            #region Call Sync command with Add element to add a no-recurrence meeting to the server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail,this.Site);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                elementsToValueMap.Add(Request.ItemsChoiceType8.ResponseRequested, true);
            }

            // Set the reminder to 10 minutes.
            string reminder = "10";
            elementsToValueMap.Add(Request.ItemsChoiceType8.Reminder, reminder);

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0", StringComparison.CurrentCultureIgnoreCase) || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1", StringComparison.CurrentCultureIgnoreCase))
            {
                Request.Location location = new Request.Location();
                location.Accuracy = (double)1;
                location.AccuracySpecified = true;
                location.Altitude = (double)55.46;
                location.AltitudeAccuracy = (double)1;
                location.AltitudeAccuracySpecified = true;
                location.AltitudeSpecified = true;
                location.Annotation = "Location sample annotation";
                location.City = "Location sample city";
                location.Country = "Location sample country";
                location.DisplayName = "Location sample dislay name";
                location.Latitude = (double)11.56;
                location.LatitudeSpecified = true;
                location.LocationUri = "Location Uri";
                location.Longitude = (double)1.9;
                location.LongitudeSpecified = true;
                location.PostalCode = "Location sample postal code";
                location.State = "Location sample state";
                location.Street = "Location sample street";
                elementsToValueMap.Add(Request.ItemsChoiceType8.Location, location);
            }

            // Call Sync command with Add element to add a meeting
            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            // Add the meeting to clean up list.
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1074");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1074
            Site.CaptureRequirementIfAreEqual<string>(
                subject,
                calendar.Calendar.Subject,
                1074,
                @"[In Sending and Receiving Meeting Requests] When a user creates a meeting on the client, the calendar item is added to the server by using the Sync command ([MS-ASCMD] section 2.2.1.21).");
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request in atendee's inbox folder and the calendar item in attendee's calendar folder
            // Get the meeting request mail.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);

            // Get the calendar item in attendee's calendar folder.
            SyncStore getCalendarResult = this.GetSyncResult(subject, this.User2Information.CalendarCollectionId, null);
            Sync calendarOfAttendee = TestSuiteHelper.GetSyncAddItem(getCalendarResult, subject);
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R17");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R17
            Site.CaptureRequirementIfAreEqual<string>(
                subject,
                calendarOfAttendee.Calendar.Subject,
                17,
                @"[In Sending and Receiving Meeting Requests] When an attendee's Calendar folder is synchronized, the Sync command response from the server contains the new calendar item that is to be added to the attendee's Calendar folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R236");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R236
            Site.CaptureRequirementIfAreEqual<byte?>(
                0,
                item.Email.MeetingRequest.AllDayEvent,
                236,
                @"[In AllDayEvent] If the value of this element is set to 0 (zero), the meeting request does not correspond to an all-day event.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R389");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R389
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                item.Email.MeetingRequest.DisallowNewTimeProposal,
                389,
                @"[In DisallowNewTimeProposal] If this element[DisallowNewTimeProposal] is not specified, the value defaults to 0 (zero), meaning that new time proposals are allowed.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R687");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R687
            ushort reminderInSeconds = (ushort)(uint.Parse(reminder) * 60);
            Site.CaptureRequirementIfAreEqual<ushort>(
                reminderInSeconds,
                item.Email.MeetingRequest.Reminder,
                687,
                @"[In Reminder] The Reminder element is an optional child element of the MeetingRequest element (section 2.2.2.48) that specifies the number of seconds prior to the calendar item's start time that a reminder will be displayed.");

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R715");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R715
                Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    item.Email.MeetingRequest.ResponseRequested,
                    715,
                    @"[In ResponseRequested] A ResponseRequested element value of 1 indicates that a response is requested.");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC03_MeetingRequest_Weekly
        /// <summary>
        /// This case is designed to test a weekly meeting request that contains DayOfWeek element and FirstDayOfWeek element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC03_MeetingRequest_Weekly()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a weekly meeting to the server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                elementsToValueMap.Add(Request.ItemsChoiceType8.ResponseRequested, false);
            }

            // Set the recurrence type to 1, which means the meeting recurs weekly.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 1,
                Interval = 1,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                Until = DateTime.UtcNow.AddDays(20).ToString("yyyyMMddTHHmmssZ")
            };

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1"))
            {
                recurrence.FirstDayOfWeek = 1;
                recurrence.FirstDayOfWeekSpecified = true;
            }

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            // Call Sync command with Add element to add a meeting
            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Get the meeting mail.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);
            #endregion

            #region Verify requirement
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R417");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R417
                Site.CaptureRequirementIfAreEqual<byte>(
                    1,
                    item.Email.MeetingRequest.Recurrences.Recurrence.FirstDayOfWeek,
                    417,
                    @"[In FirstDayOfWeek] The email2:FirstDayOfWeek element is an optional child element of the Recurrence element (section 2.2.2.60) that specifies which day is considered the first day of the calendar week for the recurrence.");
            }

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R716");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R716
                Site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    item.Email.MeetingRequest.ResponseRequested,
                    716,
                    @"[In ResponseRequested] A ResponseRequested element value of 0 (zero) indicates that a response is not requested.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R936");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R936
            Site.CaptureRequirementIfIsFalse(
                item.Email.MeetingRequest.RecurrenceIdSpecified,
                936,
                @"[In RecurrenceId] the server MUST NOT include this element if none exception to a recurring meeting.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC04_MeetingRequest_Monthly_ByDay
        /// <summary>
        /// This case is designed to test a Monthly meeting request that recurs monthly on the Nth day of the month.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC04_MeetingRequest_Monthly_ByDay()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a monthly meeting that recurs monthly on the Nth day of the month to the server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Set the DisallowNewTimeProposal to true
                elementsToValueMap.Add(Request.ItemsChoiceType8.DisallowNewTimeProposal, true);
            }

            // Set the recurrence type to 2, which means the meeting recurs monthly on the first day.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 2,
                Interval = 1,
                DayOfMonth = 2,
                DayOfMonthSpecified = true,
                Occurrences = 3,
                OccurrencesSpecified = true
            };

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Ensure the meeting request is received by attendee.
            this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            SyncStore getMeetingRequest = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getMeetingRequest, subject);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R291");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R291
                Site.CaptureRequirementIfIsNotNull(
                    emailItem.Email.MeetingRequest.Recurrences.Recurrence.CalendarType,
                    291,
                    @"[In CalendarType] This element[email2:CalendarType] is required when the Type element (section 2.2.2.80) value is 2, 3, 5, or 6 in server responses.");
            }

            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC05_MeetingRequest_Monthly_ByWeek
        /// <summary>
        /// This case is designed to test a Monthly meeting request that contains WeekOfMonth element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC05_MeetingRequest_Monthly_ByWeek()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a monthly meeting contains WeekOfMonth element to the server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Set the DisallowNewTimeProposal to true
                elementsToValueMap.Add(Request.ItemsChoiceType8.DisallowNewTimeProposal, true);
            }

            // Set the recurrence type to 3, which means the meeting recurs monthly on the first Monday.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 3,
                Interval = 1,
                WeekOfMonth = 1,
                WeekOfMonthSpecified = true,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                Occurrences = 3,
                OccurrencesSpecified = true
            };

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Get the meeting mail.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);
            #endregion

            #region Verify requirements
            Response.MeetingRequest meetingRequest = item.Email.MeetingRequest;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R16");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R16
            Site.CaptureRequirementIfAreEqual<string>(
                subject,
                item.Email.Subject,
                16,
                @"[In Sending and Receiving Meeting Requests] When an attendee's Inbox folder is synchronized, the Sync command response ([MS-ASCMD] section 2.2.1.21) from the server contains the new meeting request that is to be added to the attendee's Inbox folder.");

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R390");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R390
                Site.CaptureRequirementIfAreNotEqual<byte>(
                    0,
                    meetingRequest.DisallowNewTimeProposal,
                    390,
                    @"[In DisallowNewTimeProposal] A nonzero value indicates that new time proposals are not allowed.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R291");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R291
                Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.Recurrences.Recurrence.CalendarType,
                    291,
                    @"[In CalendarType] This element[email2:CalendarType] is required when the Type element (section 2.2.2.80) value is 2, 3, 5, or 6 in server responses.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R650");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R650
            bool isVerifiedR650 = meetingRequest.Recurrences.Recurrence != null
                && string.Equals("2", meetingRequest.Recurrences.Recurrence.DayOfWeek, StringComparison.CurrentCultureIgnoreCase);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR650,
                650,
                @"[In Recurrence] The Recurrence element is a container ([MS-ASDTYPE] section 2.2) element that defines when [and how often] the meeting recurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1067");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1067
            bool isVerifiedR1067 = meetingRequest.Recurrences.Recurrence != null
                && string.Equals("1", meetingRequest.Recurrences.Recurrence.Interval, StringComparison.CurrentCultureIgnoreCase);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1067,
                1067,
                @"[In Recurrence] The Recurrence element is a container ([MS-ASDTYPE] section 2.2) element that defines [when and] how often the meeting recurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R870");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R870
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                meetingRequest.Recurrences.Recurrence.WeekOfMonth,
                870,
                @"[In WeekOfMonth] The WeekOfMonth element is an optional child element of the Recurrence element (section 2.2.2.60) that specifies the week of the month in which the meeting recurs.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC06_MeetingRequest_Yearly_IncludedIsLeapMonth
        /// <summary>
        /// This case is designed to test a Yearly meeting request that contains DayOfMonth element, MonthOfYear element and IsLeapMonth element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC06_MeetingRequest_Yearly_IncludedIsLeapMonth()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a yearly meeting with IsLeapMonth element to the Server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            // Set the recurrence type to 5, which means the meeting recurs yearly on January 1th including embolismic (leap) month.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 5,
                Interval = 1,
                DayOfMonth = 1,
                DayOfMonthSpecified = true,
                MonthOfYear = 1,
                MonthOfYearSpecified = true,
                Occurrences = 3,
                OccurrencesSpecified = true
            };

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                recurrence.IsLeapMonth = 1;
                recurrence.IsLeapMonthSpecified = true;
            }

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Get the meeting mail.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);
            #endregion

            #region Verify requirements
            Response.MeetingRequest meetingRequest = item.Email.MeetingRequest;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R365");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R365
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                meetingRequest.Recurrences.Recurrence.DayOfMonth,
                365,
                @"[In DayOfMonth] The DayOfMonth element is an optional child element of the Recurrence element (section 2.2.2.60) that specifies the day of the month on which the meeting recurs.");

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R505");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R505
                Site.CaptureRequirementIfIsTrue(
                    meetingRequest.Recurrences.Recurrence.IsLeapMonthSpecified,
                    505,
                    @"[In IsLeapMonth] The email2:IsLeapMonth element is an optional child element of the Recurrence element (section 2.2.2.60) that specifies whether the recurrence takes place in the leap month of the given year.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R291");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R291
                Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.Recurrences.Recurrence.CalendarType,
                    291,
                    @"[In CalendarType] This element[email2:CalendarType] is required when the Type element (section 2.2.2.80) value is 2, 3, 5, or 6 in server responses.");
            }
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC07_MeetingRequest_Yearly_NotIncludedIsLeapMonth
        /// <summary>
        /// This case is designed to test a Yearly meeting request that doesn’t contain IsLeapMonth element.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC07_MeetingRequest_Yearly_NotIncludedIsLeapMonth()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a yearly meeting without IsLeapMonth element to the Server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            // Sensitivity is set to 0.
            elementsToValueMap.Add(Request.ItemsChoiceType8.Sensitivity, (byte)0);

            // Set the recurrence type to 6, which means the meeting recurs yearly on January 1th.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 6,
                Interval = 1,
                DayOfWeek = 1,
                DayOfWeekSpecified = true,
                WeekOfMonth = 2,
                WeekOfMonthSpecified = true,
                MonthOfYear = 1,
                MonthOfYearSpecified = true,
                Occurrences = 3,
                OccurrencesSpecified = true
            };

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Get the meeting mail.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);
            #endregion

            #region Verify requirements
            Response.MeetingRequest meetingRequest = item.Email.MeetingRequest;

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R509");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R509
                Site.CaptureRequirementIfAreEqual<byte>(
                    0,
                    meetingRequest.Recurrences.Recurrence.IsLeapMonth,
                    509,
                    @"[In IsLeapMonth] A default value of 0 (zero, meaning FALSE) is used if the element[email2:IsLeapMonth] value is not specified in the client request.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R291");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R291
                Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.Recurrences.Recurrence.CalendarType,
                    291,
                    @"[In CalendarType] This element[email2:CalendarType] is required when the Type element (section 2.2.2.80) value is 2, 3, 5, or 6 in server responses.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R620");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R620
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                meetingRequest.Recurrences.Recurrence.MonthOfYear,
                620,
                @"[In MonthOfYear] This element[MonthOfYear] is required when the Type element (section 2.2.2.80) is set to a value of 6, indicating that the meeting recurs yearly on the Nth day of the week of the Nth month.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R625");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R625
            Site.CaptureRequirementIfAreEqual<string>(
                "3",
                meetingRequest.Recurrences.Recurrence.Occurrences,
                625,
                @"[In Occurrences] The Occurrences element is an optional child element of the Recurrence element (section 2.2.2.60) that specifies the number of occurrences before the series of recurring meetings ends.");

            Site.Assert.IsNotNull(meetingRequest.Sensitivity, "The confidentiality level of the meeting request should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R727");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R727
            Site.CaptureRequirementIfAreEqual<string>(
                "0",
                meetingRequest.Sensitivity,
                727,
                @"[In Sensitivity] The Sensitivity element is an optional child element of the MeetingRequest element (section 2.2.2.48) that specifies the confidentiality level of the meeting request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R873");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R873
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                meetingRequest.Recurrences.Recurrence.WeekOfMonth,
                873,
                @"[In WeekOfMonth] This element[WeekOfMonth] is required when the Type element (section 2.2.2.80) value is set to 6 (indicating that the meeting recurs yearly on the Nth day of the week during the Nth month each year).");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC08_MeetingRequest_Delegate
        /// <summary>
        /// This case is designed to test the delegated meeting request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC08_MeetingRequest_Delegate()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a no recurrence meeting to the server.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User4Information.UserName, this.User4Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = resultItem.Calendar;
            calendarItem.UID = resultItem.Calendar.UID;
            this.SendMeetingRequest(subject, calendarItem);

            this.SwitchUser(this.User4Information, true);
            this.RecordCaseRelativeItems(this.User4Information.UserName, this.User4Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User4Information.UserName, this.User4Information.CalendarCollectionId, subject);
            #endregion

            #region Make sure the delegator receives the copy
            this.SwitchUser(this.User5Information, true);
            SyncStore syncItemResult;
            Sync item = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            do
            {
                Thread.Sleep(waitTime);
                SyncStore initSyncResult = this.InitializeSync(this.User5Information.InboxCollectionId);
                syncItemResult = this.SyncChanges(initSyncResult.SyncKey, this.User5Information.InboxCollectionId, null);
                if (syncItemResult != null && syncItemResult.CollectionStatus == 1)
                {
                    if (syncItemResult.AddElements != null)
                    {
                        foreach (Sync syncItem in syncItemResult.AddElements)
                        {
                            if (syncItem.Email.Subject.Contains(subject))
                            {
                                item = syncItem;
                                break;
                            }
                        }
                    }
                }

                counter++;
            }
            while ((syncItemResult == null || item == null) && counter < retryCount);

            Site.Assert.IsNotNull(
                item,
                "If the Sync command executes successfully, the item in response shouldn't be null. Retry count: {0}",
                counter);
            this.RecordCaseRelativeItems(this.User5Information.UserName, this.User5Information.InboxCollectionId, subject);

            if (Common.IsRequirementEnabled(1491, this.Site))
            {
                // Try to accept the meeting request
                // Create a meeting response request item
                Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
                {
                    UserResponse = 1,
                    CollectionId = this.User5Information.InboxCollectionId,
                    RequestId = item.ServerId
                };

                // Create a meeting response request
                MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
                MeetingResponseResponse response = this.EMAILAdapter.MeetingResponse(meetingRequest);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R539");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R539
                Site.CaptureRequirementIfAreEqual<byte>(
                    6,
                    item.Email.MeetingRequest.MeetingMessageType,
                    539,
                    @"[In MeetingMessageType] The value of email2:MeetingMessageType is 6 identifies that the meeting request has been delegated.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1491");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1491
                Site.CaptureRequirementIfAreNotEqual<int>(
                    1,
                    int.Parse(response.ResponseData.Result[0].Status),
                    1491,
                    @"[In Appendix B: Product Behavior] Implementation does support value 6 for element MeetingMessageType that indicates the meeting request MUST NOT be responded to. <1> Section 2.2.2.47:  This value 6 is supported only in Exchange 2010.");
            }
            #endregion

            #region Call Sync command to synchronize the inbox folder of the primary accout
            this.SwitchUser(this.User4Information, false);

            // Sync mailbox changes
            SyncStore syncResult = this.GetSyncResult(subject, this.User4Information.InboxCollectionId, null);
            Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncResult, subject);
            #endregion

            #region Verify requirement
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R538");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R538
            Site.CaptureRequirementIfAreEqual<byte>(
                5,
                meetingRequestEmail.Email.MeetingRequest.MeetingMessageType,
                538,
                @"[In MeetingMessageType] The value of email2:MeetingMessageType is 5 identifies the delegator's copy of the meeting request.");
            #endregion
        }

        #endregion

        #region MSASEMAIL_S04_TC09_MeetingRequest_Sender
        /// <summary>
        /// This case is designed to test the Sender element of a meeting request.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC09_MeetingRequest_Sender()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Sender element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a no recurrence meeting to the server.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = resultItem.Calendar;
            calendarItem.UID = resultItem.Calendar.UID;
            this.SendMeetingRequest(subject, calendarItem);

            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request and forward it to another user
            this.GetSyncResult(subject, this.User2Information.CalendarCollectionId, null);
            // Sync mailbox changes
            SyncStore syncResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync meetingRequest = TestSuiteHelper.GetSyncAddItem(syncResult, subject);

            // Call SmartFoward to forward the meeting request to another user
            string from = meetingRequest.Email.From;
            string to = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardContent = Common.GenerateResourceName(Site, "forward: body");

            string forwardMime = TestSuiteHelper.CreatePlainTextMime(from, to, string.Empty, string.Empty, subject, forwardContent);

            SmartForwardRequest forwardRequest = TestSuiteHelper.CreateSmartForwardRequest(this.User2Information.InboxCollectionId, meetingRequest.ServerId, forwardMime);
            SmartForwardResponse response = this.EMAILAdapter.SmartForward(forwardRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SmartForward command executes successfully.");

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.InboxCollectionId, subject);

            this.SwitchUser(this.User3Information, true);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.CalendarCollectionId, subject);
            this.RecordCaseRelativeItems(this.User3Information.UserName, this.User3Information.DeletedItemsCollectionId, subject);
            #endregion

            #region Call Sync command to get the forwarded meeting request
            this.GetSyncResult(subject, this.User3Information.CalendarCollectionId, null);
            // Sync mailbox changes
            syncResult = this.GetSyncResult(subject, this.User3Information.InboxCollectionId, null);
            Sync forwardedMeetingRequest = TestSuiteHelper.GetSyncAddItem(syncResult, subject);
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R722");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R722
            Site.CaptureRequirementIfIsNotNull(
                forwardedMeetingRequest.Email.Sender,
                722,
                @"[In Sender] This element[email2:Sender] is set by the server [and is read-only on the client].");

            // The meeting request recipient forwards the meeting request to another one, so the Sender shoud be the recipient
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R724.");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R724
            Site.CaptureRequirementIfAreEqual<string>(
                meetingRequest.Email.To,
                forwardedMeetingRequest.Email.Sender,
                724,
                @"[In Sender] If present, the email2:Sender element identifies the user that actually sent the message.");

            // The meeting request is orignally sent by the "From" address, so the "From" value in the forward meeting request should be the orignal "From".
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1073");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1073
            Site.CaptureRequirementIfAreEqual<string>(
                meetingRequest.Email.From,
                forwardedMeetingRequest.Email.From,
                1073,
                @"[In Sender] If present, the From element identifies the user on whose behalf the message was sent.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC10_ExceptionToARecurringMeeting
        /// <summary>
        /// This case is designed to test a single instance exception to a recurring meeting.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC10_ExceptionToARecurringMeeting()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a weekly meeting with exception to the Server
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            // Set the recurrence type to 1, which means the meeting recurs weekly.
            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 1,
                Interval = 1,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                Occurrences = 3,
                OccurrencesSpecified = true
            };

            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);

            DateTime tempTime = DateTime.UtcNow;
            string startTime = tempTime.ToString("yyyyMMddTHHmmssZ");
            elementsToValueMap.Add(Request.ItemsChoiceType8.StartTime, startTime);

            string endTime = tempTime.AddHours(2).ToString("yyyyMMddTHHmmssZ");
            elementsToValueMap.Add(Request.ItemsChoiceType8.EndTime, endTime);

            // Set the exceptions
            Request.Exceptions exceptions = new Request.Exceptions();
            List<Request.ExceptionsException> exceptionList = new List<Request.ExceptionsException>();
            Request.ExceptionsException exception = new Request.ExceptionsException();
            int additionalDays = 8 - tempTime.DayOfWeek.GetHashCode();
            exception.ExceptionStartTime = tempTime.AddDays(additionalDays).ToString("yyyyMMddTHHmmssZ");
            exceptionList.Add(exception);

            exceptions.Exception = exceptionList.ToArray();
            elementsToValueMap.Add(Request.ItemsChoiceType8.Exceptions, exceptions);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(subject, calendarItem);

            // Switch the current user to user2.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the meeting request
            // Ensure the meeting request is received by attendee.
            SyncStore syncChangeResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);
            #endregion

            #region Verify requirements
            byte[] globalObjIdBytes = Convert.FromBase64String(item.Email.MeetingRequest.GlobalObjId);
            int totalLength = globalObjIdBytes.Length;
            int pos = 0;

            Site.Assert.IsTrue(totalLength - pos >= 16, "GlobalObjId should have 16 bytes to store CLASSID field");
            pos += 16;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store INSTDATE field");

            string instDate = TestSuiteHelper.BytesToHex(globalObjIdBytes, pos, 4);
            int month = Convert.ToInt32(instDate.Substring(4, 2), 16);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20007");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20007
            Site.CaptureRequirementIfIsTrue(
                month >= 1 && month <= 12,
                20007,
                @"[In GlobalObjId] MONTH = %x01-12");

            int date = Convert.ToInt32(instDate.Substring(6, 2), 16);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20009");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20009
            Site.CaptureRequirementIfIsTrue(
                date >= 1 && date <= 31,
                20009,
                @"[In GlobalObjId] DATE = %x01-31");

            // If the requirements MS-ASEMAIL_R20007 and MS-ASEMAIL_R20009 can be captured successfully, it means there are 2 bytes of INSTDATE to store MONTH and DATE, and there are still 2 bytes left, then requirement MS-ASEMAIL_R20003 and MS-ASEMAIL_R20005 can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20003");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20003
            Site.CaptureRequirement(
                20003,
                @"[In GlobalObjId] YEARHIGH = BYTE");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20005");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20005
            Site.CaptureRequirement(
                20005,
                @"[In GlobalObjId] YEARLOW = BYTE");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC11_MeetingRequest_FullUpdate
        /// <summary>
        /// This case is designed to test the meeting request which is fully updated.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC11_MeetingRequest_FullUpdate()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a no recurrence meeting to the server.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);

            // Reset the UID element.
            elementsToValueMap.Remove(Request.ItemsChoiceType8.UID);
            string calendarUID = Guid.NewGuid().ToString();
            elementsToValueMap.Add(Request.ItemsChoiceType8.UID, calendarUID);

            // Set the AllDayEvent to 1
            elementsToValueMap.Add(Request.ItemsChoiceType8.AllDayEvent, (byte)1);

            // Set the reminder to 10 minutes.
            elementsToValueMap.Add(Request.ItemsChoiceType8.Reminder, "10");

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = resultItem.Calendar;
            calendarItem.UID = calendarUID;
            this.SendMeetingRequest(subject, calendarItem);
            #endregion

            #region Call Sync command to get the meeting request and accept it
            this.SwitchUser(this.User2Information, true);

            // Sync mailbox changes
            SyncStore syncChangeResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);

            // Accept the meeting request
            // Create a meeting response request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = 1,
                CollectionId = this.User2Information.InboxCollectionId,
                RequestId = meetingRequestEmail.ServerId
            };

            // Create a meeting response request
            MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse response = this.EMAILAdapter.MeetingResponse(meetingRequest);

            Site.Assert.AreEqual<int>(
                 1,
                 int.Parse(response.ResponseData.Result[0].Status),
                 "The server should return an empty xml response data to indicate MeetingResponse command success.");

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call SendMail command to send the updated meeting request to attendee
            this.SwitchUser(this.User1Information, false);
            Calendar newCalendarItem = resultItem.Calendar;
            newCalendarItem.UID = calendarUID;
            newCalendarItem.Location = "Room A";
            newCalendarItem.AllDayEvent = (byte)0;
            newCalendarItem.EndTime = DateTime.UtcNow.AddHours(2);
            this.SendMeetingRequest(subject, newCalendarItem);

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
            #endregion

            #region Call Sync command to get the updated meeting request
            this.SwitchUser(this.User2Information, false);

            // Sync to get the meeting request.
            SyncStore getMeetingResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(getMeetingResult, subject);
            #endregion

            #region Verify requirements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R535");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R535
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                emailItem.Email.MeetingRequest.MeetingMessageType,
                535,
                @"[In MeetingMessageType] The value of email2:MeetingMessageType is 2 means full update.");

            SyncStore getFirstMeetingResult = this.GetSyncResult(subject, this.User2Information.DeletedItemsCollectionId, null);
            Sync firstMeetingRequestEmail = TestSuiteHelper.GetSyncAddItem(getFirstMeetingResult, subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R531");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R531
            Site.CaptureRequirementIfAreEqual<byte>(
                meetingRequestEmail.Email.MeetingRequest.MeetingMessageType,
                firstMeetingRequestEmail.Email.MeetingRequest.MeetingMessageType,
                531,
                @"[In MeetingMessageType] The email2:MeetingMessageType value is not change tracked within e-mail messages, and therefore is not updated if the value is changed after the meeting request is sent to the client.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC12_OutdatedMeetingRequest
        /// <summary>
        /// This case is designed to test that the value of MeetingMessageType element is 4, when a new meeting request or meeting updated was received.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC12_OutdatedMeetingRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The MeetingMessageType element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a meeting to the server
            string oldEmailSubject = Common.GenerateResourceName(Site, "subject");
            string organizerEmailAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar oldCalendar = TestSuiteHelper.CreateDefaultCalendar(oldEmailSubject, organizerEmailAddress, attendeeEmailAddress);
            oldCalendar.DtStamp = DateTime.UtcNow;
            oldCalendar.UID = Guid.NewGuid().ToString();
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            // Send meeting request mail
            this.SendMeetingRequest(oldEmailSubject, oldCalendar);

            // Ensure the old meeting request email has arrived at recipient's inbox folder
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information, true);
            this.GetSyncResult(oldEmailSubject, this.User2Information.InboxCollectionId, null);
            #endregion

            #region Call Sync command with Add element to add a new meeting with the same UID with the meeting sent previoursly
            this.SwitchUser(this.User1Information, false);

            // Create one new Calendar
            string newEmailSubject = Common.GenerateResourceName(Site, "subject");
            Calendar newCalendar = TestSuiteHelper.CreateDefaultCalendar(newEmailSubject, organizerEmailAddress, attendeeEmailAddress);
            newCalendar.DtStamp = oldCalendar.DtStamp.Value.AddHours(-1);
            newCalendar.UID = oldCalendar.UID;
            #endregion

            #region Call SendMail command to send the new meeting to attendee
            // Send new meeting request mail
            this.SendMeetingRequest(newEmailSubject, newCalendar);

            // Sync changes and get email item with oldEmailSubject
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information, false);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, oldEmailSubject, newEmailSubject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, oldEmailSubject);
            #endregion

            #region Call Sync command to get the meeting request
            // Sync mailbox changes
            SyncStore syncChangeResult = this.GetSyncResult(newEmailSubject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(syncChangeResult, newEmailSubject);
            #endregion

            #region Verify requirement
            // MeetingMessageType is 4 means it is one outdated meeting request, then MS-ASEMAIL_R1070 is verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1070");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1070
            Site.CaptureRequirementIfAreEqual<byte>(
                4,
                emailItem.Email.MeetingRequest.MeetingMessageType,
                1070,
                @"[In MeetingMessageType] The value of email2:MeetingMessageType is 4 means A newer meeting request or meeting update was received after this message.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC13_GlobalObjIdFormat_vCalUID
        /// <summary>
        /// This case is designed to test the format of GlobalObjId when the GlobalObjId corresponds to a vCal-Uid.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC13_GlobalObjIdFormat_vCalUID()
        {
            bool isGlobalObjIdSupported = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                    || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0")
                    || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1");
            Site.Assume.IsTrue(isGlobalObjIdSupported, "The GlobalObjId element is only supported when the MS-ASProtocolVersion header is set to 12.1, 14.0 and 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a meeting to the server
            string calendarSubject = Common.GenerateResourceName(Site, "subject");
            string organizerAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string attendeeAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string vcalUid = Guid.NewGuid().ToString();

            DateTime startTime = Convert.ToDateTime(DateTime.Now.Date.ToString("yyyy-MM-dd HH:mm:ss")).ToUniversalTime();
            DateTime endTime = startTime.AddDays(1);
            Calendar newCalendar = this.CreateDefaultCalendar(calendarSubject, organizerAddress, attendeeAddress, vcalUid, null, startTime, endTime);

            // Record the calendar item that created in calendar folder of user1
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region The sender calls SendMail command to send to meeting request to the recipient.
            newCalendar.AllDayEvent = 1;
            this.SendMeetingRequest(calendarSubject, newCalendar);
            #endregion

            #region The recipient calls Sync method to synchronize the meeting request on the server.
            this.SwitchUser(this.User2Information, true);
            SyncStore syncChangeResult = this.GetSyncResult(calendarSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, calendarSubject);

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, calendarSubject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Verify requirements
            byte[] globalObjIdBytes = Convert.FromBase64String(item.Email.MeetingRequest.GlobalObjId);
            int totalLength = globalObjIdBytes.Length;
            int pos = 0;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R235");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R235
            Site.CaptureRequirementIfAreEqual<byte>(
                1,
                item.Email.MeetingRequest.AllDayEvent,
                235,
                @"[In AllDayEvent] If the value of this element is set to 1, the meeting request corresponds to an all-day event.");

            // The globalObjId string have been decoded into a System.Byte array by using the base64 format (See above code)
            // As we know, the value of System.Byte must be in between 0x00 (Byte.MinValue) and 0xFF (Byte.MaxValue)
            // So if reach here, this requirement can be captured directly since the meaning of 'BYTE' in techinical specification
            // is synonymous with System.Byte. That's to say, it's just a known byte in the the computer world.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20022");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20022
            Site.CaptureRequirement(
                20022,
                @"[In GlobalObjId] BYTE = %x00-FF");

            Site.Assert.IsTrue(totalLength - pos >= 16, "GlobalObjId should have 16 bytes to store CLASSID field");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20000");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20000
            Site.CaptureRequirementIfAreEqual<string>(
                "040000008200E00074C5B7101A82E008",
                TestSuiteHelper.BytesToHex(globalObjIdBytes, pos, 16),
                20000,
                @"[In GlobalObjId] CLASSID = %x04 %x00 %x00 %x00 %x82 %x00 %xE0 %x00 %x74 %xC5 %xB7 %x10 %x1A %x82 %xE0 %x08");

            pos += 16;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store INSTDATE field");

            // If the routine can reach here, then it indicates the INSTDATE equals "x00 %x00 %x00 %x00" or equals "YEARHIGH YEARLOW MONTH DATE".
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20001");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20001
            Site.CaptureRequirement(
                20001,
                @"[In GlobalObjId] INSTDATE = (%x00 %x00 %x00 %x00) | (YEARHIGH YEARLOW MONTH DATE)");

            pos += 4;
            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId should have 8 bytes to store NOW field");

            // It can be captured directly since the left have at least 8 bytes
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20011");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20011
            Site.CaptureRequirement(
                20011,
                @"[In GlobalObjId] NOW = 4BYTE 4BYTE ");

            pos += 8;
            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId should have 8 bytes to store ZERO field");
                        
            // It can be captured directly since the left have reserved 8 bytes.          
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10000");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10000                                       
            Site.CaptureRequirement(
                10000,
                @"[In GlobalObjId] Reserved bytes.RESERVED = 8BYTE");

            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId Reserved bytes.RESERVED = 8BYTE");

            pos += 8;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store BYTECOUNT field");

            int byteCount = BitConverter.ToInt32(globalObjIdBytes, pos);
            int uidLength = byteCount - 13;

            pos += 4;
            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId should have 8 bytes to store VCALSTRING field");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20018");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20018
            Site.CaptureRequirementIfAreEqual<string>(
                "vCal-Uid",
                Encoding.ASCII.GetString(globalObjIdBytes, pos, 8),
                20018,
                @"[In GlobalObjId] VCALSTRING = ""vCal-Uid""");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1001");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1001
            // If the GlobalObjId can successfully convert to vCal-Uid, it means the BYTECOUNT is really 4 bytes, then requirement MS-ASEMAIL_R1001 can be captured.
            Site.CaptureRequirement(
                1001,
                @"[In GlobalObjId] BYTECOUNT = 4BYTE");

            pos += 8;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store VERSION field");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20019");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20019
            Site.CaptureRequirementIfAreEqual<string>(
                "01000000",
                TestSuiteHelper.BytesToHex(globalObjIdBytes, pos, 4),
                20019,
                @"[In GlobalObjId] VERSION = %x01 %x00 %x00 %x00");

            pos += 4;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20021");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20021
            bool isVerifiedR20021 = totalLength - pos >= uidLength;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR20021,
                20021,
                @"[In GlobalObjId] UID = *BYTE ");

            pos += uidLength;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20016");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20016
            bool isVerifiedR20016 = (pos == globalObjIdBytes.Length - 1) && (globalObjIdBytes[pos] == 0x00);

            // If the last byte equals the 0x00, the requirement can be captured since we have
            // verified the VCALSTRING, VERSION, and UID field
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR20016,
                20016,
                @"[In GlobalObjId] VCALID = VCALSTRING VERSION UID %x00");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1002");

            //If 20016 is verified, then prove the last byte equals the 0x00, so this requirement can be captured directly
            Site.CaptureRequirement(
               1002,
               @"[In GlobalObjId] [In GlobalObjId] NULL = %x00");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20013");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20013
            // If the routine can reach here, then it indicates the DATA can meet VCALID format
            Site.CaptureRequirement(
                20013,
                @"[In GlobalObjId] DATA = OUTLOOKID | VCALID");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R121");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R121
            // If the routine can reach here, then it indicates the GlobalObjId can meet the format.
            Site.CaptureRequirement(
                121,
                @"[In GlobalObjId] The following Augmented Backus-Naur Form (ABNF) notation specifies the format of the GlobalObjId element. GLOBALOBJID =  CLASSID INSTDATE NOW RESERVED BYTECOUNT DATA");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC14_GlobalObjIdFormat_OutlookID
        /// <summary>
        /// This case is designed to test the format of GlobalObjId when the GlobalObjId corresponds to an outlook id.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC14_GlobalObjIdFormat_OutlookID()
        {
            bool isGlobalObjIdSupported = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0")
                || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1");
            Site.Assume.IsTrue(isGlobalObjIdSupported, "The GlobalObjId element is only supported when the MS-ASProtocolVersion header is set to 12.1, 14.0 and 14.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a meeting to the server
            string calendarSubject = Common.GenerateResourceName(Site, "subject");
            string organizerAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string attendeeAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            DateTime creationTime = DateTime.UtcNow;
            string outlookID = TestSuiteHelper.GenerateOutlookID(creationTime);

            Calendar newCalendar = this.CreateDefaultCalendar(calendarSubject, organizerAddress, attendeeAddress, outlookID, creationTime, null, null);

            // Record the calendar item that created in calendar folder of user1
            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region The sender calls SendMail method to send to meeting request to the recipient.
            this.SendMeetingRequest(calendarSubject, newCalendar);
            #endregion

            #region The recipient calls Sync method to synchronize the meeting request on the server.
            this.SwitchUser(this.User2Information, true);

            SyncStore syncChangeResult = this.GetSyncResult(calendarSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, calendarSubject);

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, calendarSubject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, calendarSubject);
            #endregion

            #region Verify requirement
            byte[] globalObjIdBytes = Convert.FromBase64String(item.Email.MeetingRequest.GlobalObjId);
            int totalLength = globalObjIdBytes.Length;
            int pos = 0;

            // The globalObjId string have been decoded into a System.Byte array by using the base64 format (See above code)
            // As we know, the value of System.Byte must be in between 0x00 (Byte.MinValue) and 0xFF (Byte.MaxValue)
            // So if reach here, this requirement can be captured directly since the meaning of 'BYTE' in techinical specification
            // is synonymous with System.Byte. That's to say, it's just a known byte in the the computer world.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20022");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20022
            Site.CaptureRequirement(
                20022,
                @"[In GlobalObjId] BYTE = %x00-FF");

            Site.Assert.IsTrue(totalLength - pos >= 16, "GlobalObjId should have 16 bytes to store CLASSID field");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20000");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20000
            Site.CaptureRequirementIfAreEqual<string>(
                "040000008200E00074C5B7101A82E008",
                TestSuiteHelper.BytesToHex(globalObjIdBytes, pos, 16),
                20000,
                @"[In GlobalObjId] CLASSID = %x04 %x00 %x00 %x00 %x82 %x00 %xE0 %x00 %x74 %xC5 %xB7 %x10 %x1A %x82 %xE0 %x08");

            pos += 16;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store INSTDATE field");

            // If the routine can reach here, then it indicates the INSTDATE equals "x00 %x00 %x00 %x00" or equals "YEARHIGH YEARLOW MONTH DATE"
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20001");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20001
            Site.CaptureRequirement(
                20001,
                @"[In GlobalObjId] INSTDATE = (%x00 %x00 %x00 %x00) | (YEARHIGH YEARLOW MONTH DATE)");

            pos += 4;
            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId should have 8 bytes to store NOW field");

            // It can be captured directly since the left have at least 8 bytes
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20011");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20011
            Site.CaptureRequirement(
                20011,
                @"[In GlobalObjId] NOW = 4BYTE 4BYTE ");

            pos += 8;
            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId should have 8 bytes to store ZERO field");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R10000");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R10000                                       
            Site.CaptureRequirement(
                10000,
                @"[In GlobalObjId] Reserved bytes.RESERVED = 8BYTE");

            Site.Assert.IsTrue(totalLength - pos >= 8, "GlobalObjId Reserved bytes.RESERVED = 8BYTE");

            pos += 8;
            Site.Assert.IsTrue(totalLength - pos >= 4, "GlobalObjId should have 4 bytes to store BYTECOUNT field");

            int byteCount = BitConverter.ToInt32(globalObjIdBytes, pos);
            pos += 4;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20015");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20015
            // If the last bytes count equals the actual value of BYTECOUNT field.
            // then it indicates the OUTLOOK filed really exists. Of course, it is consists of BYTE
            Site.CaptureRequirementIfAreEqual<int>(
                byteCount,
                totalLength - pos,
                20015,
                @"[In GlobalObjId] OUTLOOKID = *BYTE");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1001");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1001
            // If the GlobalObjId can successfully convert to OUTLOOKID, it means the BYTECOUNT is really 4 bytes, then requirement MS-ASEMAIL_R1001 can be captured.
            Site.CaptureRequirement(
                1001,
                @"[In GlobalObjId] BYTECOUNT = 4BYTE");

            // If the routine can reach here, then it indicates the DATA can meet OUTLOOKID
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R20013");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R20013
            Site.CaptureRequirement(
                20013,
                @"[In GlobalObjId] DATA = OUTLOOKID | VCALID");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R121");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R121
            // If the routine can reach here, then it indicates the GlobalObjId can meet the format.
            Site.CaptureRequirement(
                121,
                @"[In GlobalObjId] The following Augmented Backus-Naur Form (ABNF) notation specifies the format of the GlobalObjId element. GLOBALOBJID =  CLASSID INSTDATE NOW RESERVED BYTECOUNT DATA");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC15_RecurrenceIdExistsForException
        /// <summary>
        /// This case is designed to test server must include RecurrenceId in response messages to indicate a single instance exception to a recurring meeting.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC15_RecurrenceIdExistsForException()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The recurring calendar item cannot be created when protocol version is set to 16.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call Sync command with Add element to add a recurrence calendar
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Request.Attendees attendees = TestSuiteHelper.CreateAttendees(attendeeEmail);
            string recurrenceCalendarSubject = Common.GenerateResourceName(Site, "calendarSubject");
            string location = Common.GenerateResourceName(Site, "Room");
            DateTime currentDate = DateTime.Now.AddDays(1);
            DateTime startTime = new DateTime(currentDate.Year, currentDate.Month, currentDate.Day, 10, 0, 0);
            DateTime endTime = startTime.AddHours(10);
            string UID = Guid.NewGuid().ToString();
            string timeZone = Common.GetTimeZone("(UTC) Coordinated Universal Time", 0);

            Request.Recurrence recurrence = new Request.Recurrence
            {
                Type = 1,
                Interval = 1,
                DayOfWeek = 2,
                DayOfWeekSpecified = true,
                Occurrences = 5,
                OccurrencesSpecified = true
            };

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = new Dictionary<Request.ItemsChoiceType8, object>();
            elementsToValueMap.Add(Request.ItemsChoiceType8.Subject, recurrenceCalendarSubject);
            elementsToValueMap.Add(Request.ItemsChoiceType8.Location1, location);
            elementsToValueMap.Add(Request.ItemsChoiceType8.StartTime, startTime.ToString("yyyyMMddTHHmmssZ"));
            elementsToValueMap.Add(Request.ItemsChoiceType8.EndTime, endTime.ToString("yyyyMMddTHHmmssZ"));
            elementsToValueMap.Add(Request.ItemsChoiceType8.Recurrence, recurrence);
            elementsToValueMap.Add(Request.ItemsChoiceType8.UID, UID);
            elementsToValueMap.Add(Request.ItemsChoiceType8.MeetingStatus, (byte)1);
            elementsToValueMap.Add(Request.ItemsChoiceType8.Attendees, attendees);
            elementsToValueMap.Add(Request.ItemsChoiceType8.Timezone, timeZone);

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, recurrenceCalendarSubject);
            #endregion

            #region Call Sync command to get the added calendar item
            SyncStore getChangeResult = this.GetSyncResult(recurrenceCalendarSubject, this.User1Information.CalendarCollectionId, null);
            Sync calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, recurrenceCalendarSubject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee
            Calendar calendarItem = calendar.Calendar;
            this.SendMeetingRequest(recurrenceCalendarSubject, calendarItem);
            #endregion

            #region Call Sync command to get the meeting request and accept it
            this.SwitchUser(this.User2Information, true);

            // Sync mailbox changes
            SyncStore syncChangeResult = this.GetSyncResult(recurrenceCalendarSubject, this.User2Information.InboxCollectionId, null);
            Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncChangeResult, recurrenceCalendarSubject);

            // Accept the meeting request
            // Create a meeting response request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = 1,
                CollectionId = this.User2Information.InboxCollectionId,
                RequestId = meetingRequestEmail.ServerId
            };

            // Create a meeting response request
            MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse response = this.EMAILAdapter.MeetingResponse(meetingRequest);

            Site.Assert.AreEqual<int>(
                 1,
                 int.Parse(response.ResponseData.Result[0].Status),
                 "The MeetingResponse operation should be successful.");

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, recurrenceCalendarSubject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, recurrenceCalendarSubject);
            #endregion

            #region Call Sync command to cancel one occurrence of the meeting request
            this.SwitchUser(this.User1Information, false);
            Request.SyncCollectionChangeApplicationData changeCalednarData = new Request.SyncCollectionChangeApplicationData();
            changeCalednarData.ItemsElementName = new Request.ItemsChoiceType7[] 
            { 
                Request.ItemsChoiceType7.Location1,
                Request.ItemsChoiceType7.Recurrence,
                Request.ItemsChoiceType7.Exceptions,
                Request.ItemsChoiceType7.UID,
                Request.ItemsChoiceType7.Attendees,
                Request.ItemsChoiceType7.MeetingStatus,
                Request.ItemsChoiceType7.Subject
            };

            Request.ExceptionsException exception = new Request.ExceptionsException();
            int additionalDays = 8 - startTime.DayOfWeek.GetHashCode();
            exception.ExceptionStartTime = startTime.AddDays(additionalDays + 7).ToString("yyyyMMddTHHmmssZ");
            Request.Exceptions exceptions = new Request.Exceptions() { Exception = new Request.ExceptionsException[] { exception } };

            changeCalednarData.Items = new object[] { location, recurrence, exceptions, UID, attendees, (byte)1, recurrenceCalendarSubject };

            Request.SyncCollectionChange appDataChange = new Request.SyncCollectionChange
            {
                ApplicationData = changeCalednarData,
                ServerId = calendar.ServerId
            };

            SyncRequest syncRequest = TestSuiteHelper.CreateSyncChangeRequest(getChangeResult.SyncKey, this.User1Information.CalendarCollectionId, appDataChange);
            SyncStore result = this.EMAILAdapter.Sync(syncRequest);
            Site.Assert.AreEqual<byte>(
                1,
                result.CollectionStatus,
                "The server returns a Status 1 in the Sync command response indicate sync command success.");

            getChangeResult = this.GetSyncResult(recurrenceCalendarSubject, this.User1Information.CalendarCollectionId, null);
            calendar = TestSuiteHelper.GetSyncAddItem(getChangeResult, recurrenceCalendarSubject);

            string icalendarFormatContent = TestSuiteHelper.CreateiCalendarFormatCancelContent(calendar.Calendar);
            string meetingEmailMime = TestSuiteHelper.CreateMeetingRequestMime(
                calendar.Calendar.OrganizerEmail,
                calendar.Calendar.Attendees.Attendee[0].Email,
                recurrenceCalendarSubject,
                Common.GenerateResourceName(Site, "content"),
                icalendarFormatContent);
            string clientId = TestSuiteHelper.GetClientId();

            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, meetingEmailMime);
            SendMailResponse sendMailResponse = this.EMAILAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 sendMailResponse.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command success.");
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, recurrenceCalendarSubject);
            #endregion

            #region Call Sync command to get the udpated meeting request
            this.SwitchUser(this.User2Information, false);
            syncChangeResult = this.GetSyncResult(recurrenceCalendarSubject, this.User2Information.InboxCollectionId, null);
            Sync item = TestSuiteHelper.GetSyncAddItem(syncChangeResult, recurrenceCalendarSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R935");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R935
            Site.CaptureRequirementIfIsTrue(
                item.Email.MeetingRequest.RecurrenceIdSpecified,
                935,
                @"[In RecurrenceId] The server MUST include this element[RecurrenceId] in response messages to indicate a single instance exception to a recurring meeting.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R679");

            Site.CaptureRequirementIfAreEqual<DateTime>(
                DateTime.ParseExact(exception.ExceptionStartTime, "yyyyMMddTHHmmssZ", System.Globalization.CultureInfo.CurrentCulture).Date,
                item.Email.MeetingRequest.RecurrenceId.Date,
                679,
                @"[In RecurrenceId] The value of this element[RecurrenceId] MUST be the date corresponding to this instance of a recurring item.");

            if (Common.IsRequirementEnabled(1493, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R1493");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R1493
                Site.CaptureRequirementIfAreEqual<string>(
                    exception.ExceptionStartTime,
                    item.Email.MeetingRequest.RecurrenceId.ToString("yyyyMMddTHHmmssZ"),
                    1493,
                    @"[In RecurrenceId] The value of this element[RecurrenceId] does include the original start time of the instance if possible. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC16_BusyStatusIsNotPresent
        /// <summary>
        /// This case is designed to test if element BusyStatus is not present, a default value of 2 MUST be assumed.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC16_BusyStatusIsNotPresent()
        {
            #region Call Sync command with Add element to add a no recurrence meeting to the server.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);
            DateTime startTime = DateTime.Now.AddMinutes(-5);
            DateTime endTime = startTime.AddHours(1);
            elementsToValueMap.Add(Request.ItemsChoiceType8.StartTime, startTime.ToString("yyyyMMddTHHmmssZ"));
            elementsToValueMap.Add(Request.ItemsChoiceType8.EndTime, endTime.ToString("yyyyMMddTHHmmssZ"));

            this.SyncAddMeeting(this.User1Information.CalendarCollectionId, elementsToValueMap);

            this.RecordCaseRelativeItems(this.User1Information.UserName, this.User1Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the added calendar item.
            SyncStore getChangeResult = this.GetSyncResult(subject, this.User1Information.CalendarCollectionId, null);
            Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
            #endregion

            #region Call SendMail command to send the meeting request to attendee without setting BusyStatus.
            Calendar calendarItem = resultItem.Calendar;
            calendarItem.BusyStatus = null;
            this.SendMeetingRequest(subject, calendarItem);
            #endregion

            #region Call Sync command to get the meeting request and accept it.
            this.SwitchUser(this.User2Information, true);

            // Sync mailbox changes
            SyncStore syncChangeResult = this.GetSyncResult(subject, this.User2Information.InboxCollectionId, null);
            Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);

            // Accept the meeting request
            // Create a meeting response request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = 1,
                CollectionId = this.User2Information.InboxCollectionId,
                RequestId = meetingRequestEmail.ServerId
            };

            // Create a meeting response request
            MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse response = this.EMAILAdapter.MeetingResponse(meetingRequest);

            Site.Assert.AreEqual<int>(
                 1,
                 int.Parse(response.ResponseData.Result[0].Status),
                 "The MeetingResponse operation should be successful.");

            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.DeletedItemsCollectionId, subject);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.CalendarCollectionId, subject);
            #endregion

            #region Call Sync command to get the calendar item.
            SyncStore getCalendarItemsResult = this.GetSyncResult(subject, this.User2Information.CalendarCollectionId, null);
            Sync calendarResult = TestSuiteHelper.GetSyncAddItem(getCalendarItemsResult, subject);
            Site.Assert.IsNotNull(calendarResult.Calendar.BusyStatus, "Element BusyStatus should be present.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R270");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R270
            // BusyStatus specifies the busy status of the recipient for the meeting, once the meeting request is accepted.
            // If the BusyStatus for the calendar item of the meeting is 2, default value of 2 is used.
            Site.CaptureRequirementIfAreEqual<byte>(
                2,
                (byte)calendarResult.Calendar.BusyStatus,
                270,
                @"[In BusyStatus] If this element[BusyStatus] is not present, a default value of 2 MUST be assumed.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S04_TC17_BusyStatusIsWorkingElsewhere
        /// <summary>
        /// This case is designed to test if element BusyStatus is Working Elsewhere.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S04_TC17_BusyStatusIsWorkingElsewhere()
        {
            #region Call Sync command with Add element to add a no recurrence meeting to the server.
            string subject = Common.GenerateResourceName(Site, "Subject");
            string attendeeEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);

            Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap = TestSuiteHelper.SetMeetingProperties(subject, attendeeEmail, this.Site);
            DateTime startTime = DateTime.Now.AddMinutes(-5);
            DateTime endTime = startTime.AddHours(1);
            elementsToValueMap.Add(Request.ItemsChoiceType8.StartTime, startTime.ToString("yyyyMMddTHHmmssZ"));
            elementsToValueMap.Add(Request.ItemsChoiceType8.EndTime, endTime.ToString("yyyyMMddTHHmmssZ"));
            elementsToValueMap.Add(Request.ItemsChoiceType8.BusyStatus, (byte)4);

            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData
            {
                Items = new object[elementsToValueMap.Count],
                ItemsElementName = new Request.ItemsChoiceType8[elementsToValueMap.Count]
            };

            if (elementsToValueMap.Count > 0)
            {
                elementsToValueMap.Values.CopyTo(applicationData.Items, 0);
                elementsToValueMap.Keys.CopyTo(applicationData.ItemsElementName, 0);
            }

            SyncStore iniSync = this.InitializeSync(this.User1Information.CalendarCollectionId);
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(iniSync.SyncKey, this.User1Information.CalendarCollectionId, applicationData);

            SyncStore syncAddResponse = this.EMAILAdapter.Sync(syncAddRequest);

            #endregion
        }
        #endregion
    }
}