namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to updating of calendar related items on server.
    /// </summary>
    [TestClass]
    public class S02_UpdateCalendarRelatedItem : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is designed to test updating a single appointment item(calendar item without attendeeType) successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC01_UpdateSingleCalendarItem()
        {
            #region Organizer creates a single calendar
            #region Define a calendar item to update
            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.Location = this.Location;
            if (Common.IsRequirementEnabled(16505, this.Site))
            {
                calendarItem.LegacyFreeBusyStatus = LegacyFreeBusyType.WorkingElsewhere;
                calendarItem.LegacyFreeBusyStatusSpecified = true;
            }

            calendarItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            calendarItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            calendarItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "Create a single calendar item should be successful.");
            ItemIdType calendarId = item.Items.Items[0].ItemId;

            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar should exist.");

            if (Common.IsRequirementEnabled(16505, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16505");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16505
                this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                    LegacyFreeBusyType.WorkingElsewhere,
                    calendar.LegacyFreeBusyStatus,
                    16505,
                    @"[In Appendix C: Product Behavior] Implementation does support the LegacyFreeBusyStatus in t:CalendarItemType Complex Type which value set to ""WorkingElsewhere"" specifies the status as working outside the office. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(713, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R713");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R713
                // The type of StartTimeZone has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    calendar.StartTimeZone,
                    713,
                    @"[In Appendix C: Product Behavior] Implementation does support ""StartTimeZone"" with type ""t:TimeZoneDefinitionType ([MS-OXWSGTZ] section 2.2.4.12)"" which specifies the calendar item start time zone information. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(714, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R714");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R714
                // The type of EndTimeZone has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    calendar.EndTimeZone,
                    714,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""EndTimeZone"" with type ""t:TimeZoneDefinitionType"" which specifies the calendar item end time zone information. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Organizer updates the Location property of the created calendar item
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = calendarId;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Organizer, itemChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R650");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R650
            this.Site.CaptureRequirementIfIsNotNull(
                updatedItem,
                650,
                @"[In Messages] UpdateItemSoapIn: For each item being updated that is not a recurring calendar item, the ItemChange element MUST contain an ItemId child element ([MS-OXWSCORE] section 3.1.4.9.3.7).");

            #region Verify the Location of the calendar has been updated
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, calendarId);
            Site.Assert.IsNotNull(getItem, "The updated calendar should exist.");

            CalendarItemType updatedCalendar = getItem.Items.Items[0] as CalendarItemType;
            Site.Assert.AreEqual<string>(
                this.LocationUpdate,
                updatedCalendar.Location,
                string.Format("The Location of the updated calendar should be {0}. The actual value is {1}.", this.LocationUpdate, updatedCalendar.Location));
            #endregion
            #endregion

            #region Attendee gets the meeting request message
            MeetingRequestMessageType meetingRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", calendarItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(meetingRequest, "The update meeting request message should exist in attendee's inbox folder.");

            if (Common.IsRequirementEnabled(28505, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28505");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28505
                this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                    LegacyFreeBusyType.WorkingElsewhere,
                    meetingRequest.IntendedFreeBusyStatus,
                    28505,
                    @"[In Appendix C: Product Behavior] Implementation does support the IntendedFreeBusyStatus which value set to ""WorkingElsewhere"" specifies the status as working outside the office. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(80049, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R80049");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R80049
                this.Site.CaptureRequirementIfIsFalse(
                    meetingRequest.IsOrganizer,
                    80049,
                    "[In Appendix C: Product Behavior] Implementation does support the IsOrganizer, which specifies whether the current user is the organizer of the meeting. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(718, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R718");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R718
                // The type of StartTimeZone has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.StartTimeZone,
                    718,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""StartTimeZone"" with type ""t:TimeZoneDefinitionType ([MS-OXWSGTZ] section 2.2.4.12)"" which specifies the time zone for the start of the meeting item. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(719, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R719");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R719
                // The type of EndTimeZone has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.EndTimeZone,
                    719,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""EndTimeZone"" with type ""t:TimeZoneDefinitionType"" which specifies the time zone for the end of the meeting item. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(710, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R710");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R710
                // The type of StartTimeZoneId has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.StartTimeZoneId,
                    710,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""StartTimeZoneId"" with type ""xs:string"" which specifies the start time zone identifier. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(711, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R711");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R711
                // The type of EndTimeZoneId has been verified in the schema validation.
                this.Site.CaptureRequirementIfIsNotNull(
                    meetingRequest.EndTimeZoneId,
                    711,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""EndTimeZoneId"" with type ""xs:string"" which specifies the end time zone identifier. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Organizer cancels the meeting
            CancelCalendarItemType cancelMeetingItem = new CancelCalendarItemType();
            cancelMeetingItem.ReferenceItemId = updatedItem.Items.Items[0].ItemId;
            item = this.CreateSingleCalendarItem(Role.Organizer, cancelMeetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The meeting should be canceled successfully.");
            #endregion

            #region Organizer gets the deleted calendar item in Deleted Items folder.
            calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.deleteditems, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The canceled calendar item should exist in organizer's Deleted Items folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R740");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R740
            this.Site.CaptureRequirementIfAreEqual<int>(
                5,
                calendar.AppointmentState,
                740,
                "[In t:CalendarItemType Complex Type] [AppointmentState: Valid values include:] 5: the meeting corresponding to the organizer's calendar item has been cancelled");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1052");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1052
            this.Site.CaptureRequirementIfAreEqual<int>(
                5,
                calendar.AppointmentState,
                1052,
                "[In t:CalendarItemType Complex Type] [AppointmentState: Valid values include:] This value [5] is found in the Deleted Items folder of the organizer.");

            #endregion

            #region Attendee verifies canceled meeting
            CalendarItemType canceledCalendar = null;
            int counter = 0;
            while (counter < this.UpperBound)
            {
                System.Threading.Thread.Sleep(this.WaitTime);
                canceledCalendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendar.UID) as CalendarItemType;

                if (canceledCalendar.AppointmentStateSpecified && canceledCalendar.AppointmentState == 7)
                {
                    break;
                }

                counter++;
            }

            if (counter == this.UpperBound && canceledCalendar.AppointmentState != 7)
            {
                Site.Assert.Fail("Attendee should get the calendar cancelled by the organizer.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R741");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R741
            this.Site.CaptureRequirementIfAreEqual<int>(
                7,
                canceledCalendar.AppointmentState,
                741,
                "[In t:CalendarItemType Complex Type] [AppointmentState: Valid values include:] 7: the meeting corresponding to the attendee's calendar item has been cancelled");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R730");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R730
            this.Site.CaptureRequirementIfIsTrue(
                canceledCalendar.IsCancelled,
                730,
                "[In t:CalendarItemType Complex Type] [IsCancelled is] True if a meeting has been canceled.");
            #endregion

            #region Clean up organizer's sentitems and deleteditems folder, and attendee's inbox, calendar and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test updating a single meeting item successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC02_UpdateSingleMeeting()
        {
            #region Define a meeting
            CalendarItemType meetingItem = new CalendarItemType();
            int timeInterval = this.TimeInterval;
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = DateTime.Now.AddHours(timeInterval);
            meetingItem.StartSpecified = true;
            timeInterval++;
            meetingItem.End = DateTime.Now.AddHours(timeInterval);
            meetingItem.EndSpecified = true;
            meetingItem.Location = this.Location;
            meetingItem.LegacyFreeBusyStatus = LegacyFreeBusyType.NoData;
            meetingItem.LegacyFreeBusyStatusSpecified = true;

            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            #region Organizer creates a meeting with CalendarItemCreateOrDeleteOperationType value set to SendToNone
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "Create a meeting item should be successful.");
            ItemIdType meetingId = item.Items.Items[0].ItemId;

            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar should be created successfully.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16502");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16502
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.NoData,
                calendar.LegacyFreeBusyStatus,
                16502,
                @"[In t:CalendarItemType Complex Type] The LegacyFreeBusyStatus which value is ""NoData"" specifies that there is no data for that recipient.");

            #endregion

            #region Organizer updates the Location element in the created meeting item with CalendarItemUpdateOperationType value set to SendOnlyToAll
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = meetingId;

            // Update the created calendar item.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Organizer, itemChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(updatedItem, "Update the meeting item should be successful.");
            #endregion

            #region Attendee gets the meeting request message to check whether the Location element of meeting request message is updated
            MeetingRequestMessageType meetingRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(meetingRequest, "The meeting request should exist.");
            Site.Assert.AreEqual<string>(this.LocationUpdate, meetingRequest.Location, "Location in meeting request message should be updated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28502");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28502
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.NoData,
                meetingRequest.IntendedFreeBusyStatus,
                28502,
                @"[In t:MeetingRequestMessageType Complex Type] The IntendedFreeBusyStatus which value is ""NoData"" specifies that there is no data for that recipient.");
            #endregion

            #region Clean up organizer's calendar and deleteditems folders, and attendee's inbox and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test updating a recurring calendar item successfully. 
        /// It also verifies ModifiedOccurrences element in CalendarItemType and requirements related to that calendar related items can be generated by CreateItem operation.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC03_UpdateRecurringCalendar()
        {
            #region Define a recurring calendar item
            DateTime startTime = DateTime.Now;

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = this.PatternInterval;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = this.NumberOfOccurrences;
            range.StartDate = startTime;

            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = startTime;
            meetingItem.StartSpecified = true;
            meetingItem.End = startTime.AddHours(this.TimeInterval);
            meetingItem.EndSpecified = true;
            meetingItem.Location = this.Location;
            meetingItem.Recurrence = new RecurrenceType();
            meetingItem.Recurrence.Item = pattern;
            meetingItem.Recurrence.Item1 = range;
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            #region Organizer creates a recurring calendar item with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R494");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R494
            Site.CaptureRequirementIfIsNotNull(
                 item,
                 494,
                 @"[In CreateItem Operation] This operation [CreateItem] can be used to create meetings.");
            #endregion

            #region Organizer updates the Location element of the first occurrence of the recurring calendar item
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The meeting should be found in the organizer's Calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll.");

            OccurrenceItemIdType occurrenceId = new OccurrenceItemIdType();
            occurrenceId.RecurringMasterId = calendar.ItemId.Id;
            occurrenceId.InstanceIndex = 1;

            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, occurrenceId);
            Site.Assert.IsNotNull(getItem, "The updated occurrence should exist.");

            // Update the first occurrence of the recurring meeting
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = getItem.Items.Items[0].ItemId;

            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Organizer, itemChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(updatedItem, "The UpdateItem operation should be successful.");
            #endregion

            #region Attendee verifies whether the location of the first occurrence of the recurring calendar item has been updated
            CalendarItemType updatedOccurrence = null;
            bool locationUpdateSuccess = false;
            int counter = 0;
            while (counter < this.UpperBound)
            {
                System.Threading.Thread.Sleep(this.WaitTime);
                calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
                Site.Assert.IsNotNull(calendar, "The meeting should be found in attendee's Calendar folder after organizer updates the meeting.");

                occurrenceId = new OccurrenceItemIdType();
                occurrenceId.RecurringMasterId = calendar.ItemId.Id;
                occurrenceId.InstanceIndex = 1;

                getItem = this.GetSingleCalendarItem(Role.Attendee, occurrenceId);
                Site.Assert.IsNotNull(getItem, "The updated occurrence should exist.");

                updatedOccurrence = getItem.Items.Items[0] as CalendarItemType;

                locationUpdateSuccess = string.Compare(this.LocationUpdate, updatedOccurrence.Location, true) == 0;
                if (locationUpdateSuccess)
                {
                    break;
                }

                counter++;
            }

            if (counter == this.UpperBound && !locationUpdateSuccess)
            {
                Site.Assert.Fail("Attendee should get the update information of calendar by the organizer.");
            }

            MeetingRequestMessageType updateRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(updateRequest, "The meeting update request should be found in attendee's inbox folder after organizer updates the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R277");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R277
            this.Site.CaptureRequirementIfIsNotNull(
                updateRequest.RecurrenceId,
                277,
                "[In t:MeetingMessageType Complex Type] RecurrenceId: Identifies a specific instance of a recurring calendar item.");
            #endregion

            #region Attendee declines the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            DeclineItemType declineItem = new DeclineItemType();
            declineItem.ReferenceItemId = new ItemIdType();
            declineItem.ReferenceItemId = updatedOccurrence.ItemId;

            item = this.CreateSingleCalendarItem(Role.Attendee, declineItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The response messages returned by the CreateItem operation should succeed.");
            #endregion

            #region Organizer calls FindItem to re-obtain the ItemId of the updated recurring calendar item because the changekey of ItemId is updated
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The declined meeting response returned by Attendee should be in organizer's Inbox folder.");

            if (Common.IsRequirementEnabled(80049, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R80049");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R80049
                this.Site.CaptureRequirementIfIsTrue(
                    response.IsOrganizer,
                    80049,
                    "[In Appendix C: Product Behavior] Implementation does support the IsOrganizer, which specifies whether the current user is the organizer of the meeting. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(909, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R909");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R909
                Site.CaptureRequirementIfIsNotNull(
                     response.Recurrence,
                     909,
                     @"[In Appendix C: Product Behavior] Implementation does support Recurrence which is a RecurrenceType element that represents the recurrence for the calendar item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Clean up organizer's inbox, calendar and deleteditems folders, and attendee's inbox, calendar and sentitems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemUpdateOperationType set to SendOnlyToAll. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC04_UpdateMeetingWithSendOnlyToAll()
        {
            // Verify CalendarItemUpdateOperationType set to SendOnlyToAll when it used in UpdateItem operation.
            this.VerifyCalendarItemUpdateOperationType(CalendarItemUpdateOperationType.SendOnlyToAll);

            #region Clean up organizer's calendar and deleteditems folders, and attendee's inbox and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemUpdateOperationType set to SendToAllAndSaveCopy. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC05_UpdateMeetingWithSendToAllAndSaveCopy()
        {
            // Verify CalendarItemUpdateOperationType set to SendToAllAndSaveCopy when it used in UpdateItem operation.
            this.VerifyCalendarItemUpdateOperationType(CalendarItemUpdateOperationType.SendToAllAndSaveCopy);

            #region Clean up organizer's calendar, sentitems and deleteditems folders, and attendee's inbox and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemUpdateOperationType set to SendToNone. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC06_UpdateMeetingWithSendToNone()
        {
            // Verify CalendarItemUpdateOperationType set to SendToNone when it used in UpdateItem operation.
            this.VerifyCalendarItemUpdateOperationType(CalendarItemUpdateOperationType.SendToNone);

            #region Clean up organizer's calendar folder.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemUpdateOperationType set to SendOnlyToChanged. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC07_UpdateMeetingWithSendOnlyToChanged()
        {
            // Verify CalendarItemUpdateOperationType set to SendOnlyToChanged when it used in UpdateItem operation.
            this.VerifyChangeAttendeesWithCalendarItemUpdateOperationType(CalendarItemUpdateOperationType.SendOnlyToChanged);

            #region Clean up organizer's calendar folder, and attendee's inbox and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemUpdateOperationType set to SendToChangedAndSaveCopy. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC08_UpdateMeetingWithSendToChangedAndSaveCopy()
        {
            // Verify CalendarItemUpdateOperationType set to SendToChangedAndSaveCopy when it used in UpdateItem operation.
            this.VerifyChangeAttendeesWithCalendarItemUpdateOperationType(CalendarItemUpdateOperationType.SendToChangedAndSaveCopy);

            #region Clean up organizer's calendar and sentitems folders, and attendee's inbox and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test if the meeting request is an updated meeting request and the attendee has not yet responded to the 
        /// original meeting request, FullUpdate will be returned for MeetingRequestTypeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC09_MeetingRequestTypeFullUpdate()
        {
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });

            #region Organizer creates the meeting and sends it to attendee.
            CalendarItemType meeting = new CalendarItemType();
            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meeting.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            meeting.Subject = this.Subject;
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Location = this.Location;
            meeting.IsResponseRequested = true;
            meeting.IsResponseRequestedSpecified = true;
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The meeting should be created successfully.");
            #endregion

            #region Organizer updates the meeting.
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar should be created successfully.");

            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;
            AdapterHelper locationChangeInfo = new AdapterHelper();
            locationChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            locationChangeInfo.Item = new CalendarItemType() { Location = this.LocationUpdate };
            locationChangeInfo.ItemId = calendar.ItemId;
            UpdateItemResponseMessageType itemOfLocationUpdate = this.UpdateSingleCalendarItem(Role.Organizer, locationChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(itemOfLocationUpdate, "Update the meeting item should be successful.");
            #endregion

            #region Attendee gets the meeting request.
            MeetingRequestMessageType meetingRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, this.LocationUpdate, meeting.UID, UnindexedFieldURIType.calendarLocation) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(meetingRequest, "The meeting request should exist.");
            Site.Assert.AreEqual<string>(this.LocationUpdate, meetingRequest.Location, "Location in meeting request message should be updated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R484");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R484
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.FullUpdate,
                meetingRequest.MeetingRequestType,
                484,
                @"[In t:MeetingRequestTypeType Simple Type] FullUpdate: Identifies the meeting request as an updated meeting request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R485");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R485
            // Attendee does not response the meeting request and FullUpdate is returned, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.FullUpdate,
                meetingRequest.MeetingRequestType,
                485,
                @"[In t:MeetingRequestTypeType Simple Type] This value [FullUpdate] indicates that the attendee has not yet responded to the original meeting request.");
            #endregion

            #region Clean up organizer's calendar folder, and attendee's inbox, calendar and deleted items folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test if the meeting request is an updated meeting request and the attendee had previously accepted or 
        /// tentatively accepted the original meeting request, InformationalUpdate will be returned for MeetingRequestTypeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC10_MeetingRequestTypeInformationalUpdate()
        {
            #region Organizer creates the meeting and sends it to attendee.
            CalendarItemType meeting = new CalendarItemType();
            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meeting.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            meeting.Subject = this.Subject;
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Location = this.Location;
            meeting.IsResponseRequested = true;
            meeting.IsResponseRequestedSpecified = true;
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The meeting should be created successfully.");
            #endregion

            #region Attendee accepts the meeting request.
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "The meeting request message should be found in attendee's Inbox folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            AcceptItemType acceptItem = new AcceptItemType();
            acceptItem.ReferenceItemId = new ItemIdType();
            acceptItem.ReferenceItemId.Id = request.ItemId.Id;
            item = this.CreateSingleCalendarItem(Role.Attendee, acceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Accept the meeting request should be successful.");
            #endregion

            #region Organizer updates the meeting.
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar should be created successfully.");

            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;
            AdapterHelper locationChangeInfo = new AdapterHelper();
            locationChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            locationChangeInfo.Item = new CalendarItemType() { Location = this.LocationUpdate };
            locationChangeInfo.ItemId = calendar.ItemId;
            UpdateItemResponseMessageType itemOfLocationUpdate = this.UpdateSingleCalendarItem(Role.Organizer, locationChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(itemOfLocationUpdate, "Update the meeting item should be successful.");
            #endregion

            #region Attendee gets the meeting request.
            MeetingRequestMessageType meetingRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(meetingRequest, "The meeting request should exist.");
            Site.Assert.AreEqual<string>(this.LocationUpdate, meetingRequest.Location, "Location in meeting request message should be updated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R71");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R71
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.InformationalUpdate,
                meetingRequest.MeetingRequestType,
                71,
                @"[In t:MeetingRequestTypeType Simple Type] InformationUpdate: Identifies the meeting request as an updated meeting request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R587");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R587
            // Attendee has accepted the meeting and InformationalUpdate is returned, this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.InformationalUpdate,
                meetingRequest.MeetingRequestType,
                587,
                @"[In t:MeetingRequestTypeType Simple Type] This value [InformationUpdate] indicates that the attendee had previously accepted or tentatively accepted the original meeting request.");
            #endregion

            #region Clean up organizer's inbox and calendar folders, and attendee's deleted items and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.deleteditems, DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test ErrorMessageDispositionRequired will be returned if create or update a MessageType object
        /// without setting MessageDisposition property.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC11_UpdateItemErrorMessageDispositionRequired()
        {
            #region Create a message without setting MessageDisposition element.
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { new MessageType() };
            createItemRequest.Items.Items[0].Subject = this.Subject;
            DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
            folderIdForCreateItems.Id = DistinguishedFolderIdNameType.drafts;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
            CreateItemResponseType createItemResponse = this.MTGSAdapter.CreateItem(createItemRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1241");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1241
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createItemResponse.ResponseMessages.Items[0].ResponseClass,
                1241,
                @"[In Messages] ErrorMessageDispositionRequired:This error code MUST be returned under the following conditions: 
                  When the item that is being created or updated is a MessageType object. 
                  [For the CancelCalendarItemType, AcceptItemType, DeclineItemType, or TentativelyAcceptItemType response objects.]");
            #endregion

            #region Create a message with setting MessageDisposition element.
            createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            createItemRequest.MessageDispositionSpecified = true;
            createItemResponse = this.MTGSAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Update the message without setting MessageDisposition element.
            MessageType messageUpdate = new MessageType();
            messageUpdate.Subject = this.SubjectUpdate;

            UpdateItemType updateItemRequest = new UpdateItemType();
            updateItemRequest.ItemChanges = new ItemChangeType[1];
            PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
            pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemSubject;
            SetItemFieldType setItemField = new SetItemFieldType();
            setItemField.Item = pathToUnindexedField;
            setItemField.Item1 = messageUpdate;
            ItemChangeType itemChange = new ItemChangeType();
            itemChange.Item = (createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType).Items.Items[0].ItemId;
            itemChange.Updates = new ItemChangeDescriptionType[] { setItemField };
            updateItemRequest.ItemChanges[0] = itemChange;
            UpdateItemResponseType updateItemResponse = this.MTGSAdapter.UpdateItem(updateItemRequest);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1237");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1237
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                updateItemResponse.ResponseMessages.Items[0].ResponseClass,
                1237,
                @"[In Messages] If the request is unsuccessful, the UpdateItem operation returns an UpdateItemResponse element with the ResponseClass attribute of the UpdateItemResponseMessage element set to ""Error"". ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1240");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1240
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorMessageDispositionRequired,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                1240,
                @"[In Messages] ErrorMessageDispositionRequired: Occurs if the MessageDisposition property is not set. ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1241");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1241
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorMessageDispositionRequired,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                1241,
                @"[In Messages] ErrorMessageDispositionRequired:This error code MUST be returned under the following conditions: 
                  When the item that is being created or updated is a MessageType object. 
                  [For the CancelCalendarItemType, AcceptItemType, DeclineItemType, or TentativelyAcceptItemType response objects.]");
            #endregion

            #region Clean up organizer's drafts folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.drafts });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test ErrorCalendarDurationIsTooLong will be returned if duration of a calendar item exceeds five years.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S02_TC12_UpdateItemErrorCalendarDurationIsTooLong()
        {
            #region Define a calendar item
            int timeInterval = this.TimeInterval;
            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.Start = DateTime.Now.AddHours(timeInterval);
            calendarItem.StartSpecified = true;
            calendarItem.End = calendarItem.Start.AddDays(6);
            calendarItem.EndSpecified = true;
            #endregion

            #region Create the recurring calendar item and extract the Id of an occurrence item
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { calendarItem };
            createItemRequest.MessageDispositionSpecified = true;
            createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            CreateItemResponseType response = this.MTGSAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(response, 1, this.Site);
            #endregion

            #region Update the calendar to make the duration exceeds 5 years.
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.End = calendarItem.Start.AddYears(6);
            calendarUpdate.EndSpecified = true;

            UpdateItemType updateItemRequest = new UpdateItemType();
            updateItemRequest.ItemChanges = new ItemChangeType[1];
            PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
            pathToUnindexedField.FieldURI = UnindexedFieldURIType.calendarEnd;
            SetItemFieldType setItemField = new SetItemFieldType();
            setItemField.Item = pathToUnindexedField;
            setItemField.Item1 = calendarUpdate;
            ItemChangeType itemChange = new ItemChangeType();
            itemChange.Item = (response.ResponseMessages.Items[0] as ItemInfoResponseMessageType).Items.Items[0].ItemId;
            itemChange.Updates = new ItemChangeDescriptionType[] { setItemField };
            updateItemRequest.ItemChanges[0] = itemChange;
            updateItemRequest.SendMeetingInvitationsOrCancellationsSpecified = true;
            updateItemRequest.SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType.SendToNone;
            UpdateItemResponseType updateItemResponse = this.MTGSAdapter.UpdateItem(updateItemRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1244");

            // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1244
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCalendarDurationIsTooLong,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                1244,
                @"[In Messages] ErrorCalendarDurationIsTooLong: Specifies that the item duration of a calendar item exceeds five years.");
            #endregion

            #region Clean up organizer's calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Verify the value of CalendarItemUpdateOperationType.
        /// </summary>
        /// <param name="calendarItemUpdateOperationType">Specify a value of CalendarItemUpdateOperationType.</param>
        protected void VerifyCalendarItemUpdateOperationType(CalendarItemUpdateOperationType calendarItemUpdateOperationType)
        {
            #region Step1: Organizer set the properties of the meeting to create
            CalendarItemType meeting = new CalendarItemType();
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            meeting.Location = this.Location;

            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            #endregion

            #region Step2: Organizer create the meeting and sends to none
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "Create a meeting item should be successful.");
            ItemIdType meetingId = item.Items.Items[0].ItemId;
            #endregion

            #region Step3: Organizer updates a meeting with CalendarItemUpdateOperationType
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.Location = this.LocationUpdate;

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = meetingId;

            // Update the calendar item created.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Organizer, itemChangeInfo, calendarItemUpdateOperationType);
            Site.Assert.IsNotNull(updatedItem, "Update the Location of the calendar item should be successful.");

            // Get the UpdateItem and verify if the update operation is successful
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, updatedItem.Items.Items[0].ItemId);
            Site.Assert.IsNotNull(getItem, "The updated calendar item should exist.");

            bool isLocationUpdatedSuccess = ((CalendarItemType)getItem.Items.Items[0]).Location == this.LocationUpdate;
            switch (calendarItemUpdateOperationType)
            {
                case CalendarItemUpdateOperationType.SendToNone:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R60");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R60
                    this.Site.CaptureRequirementIfIsTrue(
                        isLocationUpdatedSuccess,
                        60,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToNone: The calendar item is updated.");
                    break;
                case CalendarItemUpdateOperationType.SendOnlyToAll:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R61");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R61
                    this.Site.CaptureRequirementIfIsTrue(
                        isLocationUpdatedSuccess,
                        61,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendOnlyToAll: The calendar item is updated.");
                    break;
                case CalendarItemUpdateOperationType.SendToAllAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R64");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R64
                    this.Site.CaptureRequirementIfIsTrue(
                        isLocationUpdatedSuccess,
                        64,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToAllAndSaveCopy: The calendar item is updated.");
                    break;
            }
            #endregion

            #region Step4: Verify CalendarItemUpdateOperationType used in UpdateItem operation
            #region Find the update meeting request in Organizer SentItems
            bool updatedIsFoundInOrganizerSentItems = false;
            MeetingRequestMessageType meetingRequestMessageInOrganizer = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            if (null != meetingRequestMessageInOrganizer)
            {
                if (null != meetingRequestMessageInOrganizer.Location && meetingRequestMessageInOrganizer.Location == this.LocationUpdate)
                {
                    updatedIsFoundInOrganizerSentItems = true;
                }
            }
            #endregion

            #region Find the update meeting request in Attendees Inbox
            bool updatedIsFoundInAttendeeInbox = false;
            MeetingRequestMessageType meetingRequestMessageInAttendee = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            if (null != meetingRequestMessageInAttendee)
            {
                if (null != meetingRequestMessageInAttendee.Location && meetingRequestMessageInAttendee.Location == this.LocationUpdate)
                {
                    updatedIsFoundInAttendeeInbox = true;
                }
            }
            #endregion

            switch (calendarItemUpdateOperationType)
            {
                case CalendarItemUpdateOperationType.SendToNone:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6000");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6000
                    this.Site.CaptureRequirementIfIsFalse(
                        updatedIsFoundInAttendeeInbox,
                        6000,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToNone: [The calendar item is updated] but updates are not sent to attendees.");
                    break;
                case CalendarItemUpdateOperationType.SendOnlyToAll:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6100");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6100
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInAttendeeInbox,
                        6100,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendOnlyToAll: the meeting update is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6101");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6101
                    this.Site.CaptureRequirementIfIsFalse(
                        updatedIsFoundInOrganizerSentItems,
                        6101,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendOnlyToAll: [The calendar item is updated and the meeting update is sent to all attendees] but is not saved in the folder that is specified in the request.");
                    break;
                case CalendarItemUpdateOperationType.SendToAllAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6400");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6400
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInAttendeeInbox,
                        6400,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToAllAndSaveCopy: [The calendar item is updated,] the meeting update is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6401");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6401
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInOrganizerSentItems,
                        6401,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToAllAndSaveCopy: [The calendar item is updated, the meeting update is sent to all attendees,] and a copy of the updated meeting request is saved in the Sent Items folder.");
                    break;
            }
            #endregion
        }

        /// <summary>
        /// Verify the "SendOnlyToChanged" and "SendToChangedAndSaveCopy" value of CalendarItemUpdateOperationType.
        /// </summary>
        /// <param name="calendarItemUpdateOperationType">Specify a value of CalendarItemUpdateOperationType.</param>
        protected void VerifyChangeAttendeesWithCalendarItemUpdateOperationType(CalendarItemUpdateOperationType calendarItemUpdateOperationType)
        {
            #region Step1: Organizer set the properties of the meeting to create
            CalendarItemType meeting = new CalendarItemType();
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            meeting.Location = this.Location;

            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            #endregion

            #region Step2: Organizer create the meeting and sends to none
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "Create a meeting item should be successful.");
            ItemIdType meetingId = item.Items.Items[0].ItemId;
            #endregion

            #region Step3: Organizer updates RequiredAttendees of the a meeting with CalendarItemUpdateOperationType
            CalendarItemType calendarUpdate = new CalendarItemType();
            calendarUpdate.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };

            AdapterHelper itemChangeInfo = new AdapterHelper();
            itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarRequiredAttendees;
            itemChangeInfo.Item = calendarUpdate;
            itemChangeInfo.ItemId = meetingId;

            // Update the calendar item created.
            UpdateItemResponseMessageType updatedItem = this.UpdateSingleCalendarItem(Role.Organizer, itemChangeInfo, calendarItemUpdateOperationType);
            Site.Assert.IsNotNull(updatedItem, @"Update the RequiredAttendees of the calendar item should be successful.");

            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, updatedItem.Items.Items[0].ItemId);
            Site.Assert.IsNotNull(getItem, "The updated item should exist.");
            bool isRequiredAttendeesUpdatedSuccess = string.Equals(((CalendarItemType)getItem.Items.Items[0]).RequiredAttendees[0].Mailbox.EmailAddress, this.AttendeeEmailAddress, StringComparison.OrdinalIgnoreCase);

            switch (calendarItemUpdateOperationType)
            {
                case CalendarItemUpdateOperationType.SendOnlyToChanged:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R62");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R62
                    this.Site.CaptureRequirementIfIsTrue(
                        isRequiredAttendeesUpdatedSuccess,
                        62,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendOnlyToChanged: The calendar item is updated.");
                    break;
                case CalendarItemUpdateOperationType.SendToChangedAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R65");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R65
                    this.Site.CaptureRequirementIfIsTrue(
                        isRequiredAttendeesUpdatedSuccess,
                        65,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToChangedAndSaveCopy: The calendar item is updated.");
                    break;
            }
            #endregion

            #region Step4: Verify CalendarItemUpdateOperationType used in UpdateItem operation
            #region Find the update meeting request in Organizer SentItems
            bool updatedIsFoundInOrganizerSentItems = false;
            MeetingRequestMessageType meetingRequestMessageInOrgnizer = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            if (null != meetingRequestMessageInOrgnizer)
            {
                if (null != meetingRequestMessageInOrgnizer.RequiredAttendees && meetingRequestMessageInOrgnizer.RequiredAttendees.Length > 0)
                {
                    for (int i = 0; i < meetingRequestMessageInOrgnizer.RequiredAttendees.Length; i++)
                    {
                        if (string.Equals(meetingRequestMessageInOrgnizer.RequiredAttendees[i].Mailbox.EmailAddress, this.AttendeeEmailAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            updatedIsFoundInOrganizerSentItems = true;
                        }
                    }
                }
            }

            #endregion

            #region Find the update meeting request in Attendees Inbox
            bool updatedIsFoundInAttendeeInbox = false;
            MeetingRequestMessageType meetingRequestMessageInAttendee = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            if (null != meetingRequestMessageInAttendee)
            {
                if (null != meetingRequestMessageInAttendee.RequiredAttendees && meetingRequestMessageInAttendee.RequiredAttendees.Length > 0)
                {
                    for (int i = 0; i < meetingRequestMessageInAttendee.RequiredAttendees.Length; i++)
                    {
                        if (string.Equals(meetingRequestMessageInAttendee.RequiredAttendees[i].Mailbox.EmailAddress, this.AttendeeEmailAddress, StringComparison.OrdinalIgnoreCase))
                        {
                            updatedIsFoundInAttendeeInbox = true;
                        }
                    }
                }
            }
            #endregion

            switch (calendarItemUpdateOperationType)
            {
                case CalendarItemUpdateOperationType.SendOnlyToChanged:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6200");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6200
                    // Server always treat the Organizer as RequiredAttendees, therefore, after updated the RequiredAttendees to Attendee, only Attendee will receive the updated meeting message.
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInAttendeeInbox,
                        6200,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendOnlyToChanged: [The calendar item is updated] and the meeting update is sent only to attendees that were added and/or deleted because of the update.");
                    break;
                case CalendarItemUpdateOperationType.SendToChangedAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6500");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6500
                    // Server always treat the Organizer as RequiredAttendees, therefore, after updated the RequiredAttendees to Attendee, only Attendee will receive the updated meeting message.            
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInAttendeeInbox,
                        6500,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToChangedAndSaveCopy: [The calendar item is updated,] the meeting update is sent to all attendees that were added and/or deleted as a result of the update.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R6501");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R6501
                    this.Site.CaptureRequirementIfIsTrue(
                        updatedIsFoundInOrganizerSentItems,
                        6501,
                        @"[In t:CalendarItemUpdateOperationType Simple Type] SendToChangedAndSaveCopy: [The calendar item is updated, the meeting update is sent to all attendees that were added and/or deleted as a result of the update,] and a copy of the updated meeting request is saved in the Sent Items folder.");
                    break;
            }
            #endregion
        }
        #endregion
    }
}