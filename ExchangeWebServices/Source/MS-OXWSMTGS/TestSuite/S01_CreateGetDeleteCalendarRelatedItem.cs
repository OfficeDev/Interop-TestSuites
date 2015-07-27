//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving and deletion of calendar related items on server.
    /// </summary>
    [TestClass]
    public class S01_CreateGetDeleteCalendarRelatedItem : TestSuiteBase
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
        /// This test case is designed to test getting a single calendar item with all optional elements which are empty successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC01_CreateGetDeleteSingleCalendarItem()
        {
            #region Define a calendar item
            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.ConferenceType = 0;
            calendarItem.ConferenceTypeSpecified = true;
            calendarItem.AllowNewTimeProposal = false;
            calendarItem.AllowNewTimeProposalSpecified = true;
            calendarItem.IsOnlineMeeting = false;
            calendarItem.IsOnlineMeetingSpecified = true;
            calendarItem.IsAllDayEvent = true;
            calendarItem.IsAllDayEventSpecified = true;
            calendarItem.LegacyFreeBusyStatus = LegacyFreeBusyType.OOF;
            calendarItem.LegacyFreeBusyStatusSpecified = true;
            #endregion

            #region Organizer creates the single calendar item
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R488");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R488
            Site.CaptureRequirementIfIsNotNull(
                 item,
                 488,
                 @"[In CreateItem Operation] This operation [CreateItem] can be used to create meeting request messages.");

            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The created calendar item should be found in Organizer's calendar folder.");
            ItemIdType deletedItem = calendar.ItemId;

            #region Capture Code
            bool isChecked = false;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R593");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R593
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Appointment",
                calendar.ItemClass,
                593,
                @"[In CreateItem Operation] This operation [CreateItem] can be used to create appointments.");

            Site.Assert.IsTrue(calendar.IsAllDayEventSpecified, "The value of the IsAllDayEventSpecified element should be true.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R720");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R720
            this.Site.CaptureRequirementIfIsTrue(
                calendar.IsAllDayEvent,
                720,
                @"[In t:CalendarItemType Complex Type] [IsAllDayEvent is] True if a calendar item or meeting request represents an all-day event.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R747");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R747
            isChecked = calendar.IsOnlineMeetingSpecified && !calendar.IsOnlineMeeting;
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                747,
                @"[In t:CalendarItemType Complex Type] otherwise [if the meeting is not online], [IsOnlineMeeting is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R745");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R745
            isChecked = calendar.AllowNewTimeProposalSpecified && !calendar.AllowNewTimeProposal;
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                745,
                @"[In t:CalendarItemType Complex Type] otherwise [if a new meeting time can not be proposed for a meeting by an attendee], [AllowNewTimeProposal is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R516");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R516
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                calendar.ConferenceType,
                516,
                @"[In t:CalendarItemType Complex Type] ConferenceType: Valid values include:0 (zero): video conference");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R514");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R514
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                calendar.AppointmentState,
                514,
                @"[In t:CalendarItemType Complex Type] AppointmentState: Valid values include:0 (zero): the calendar item represents an appointment");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R735");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R735
            isChecked = calendar.MeetingRequestWasSentSpecified && !calendar.MeetingRequestWasSent;
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                735,
                @"[In t:CalendarItemType Complex Type] otherwise [if request has not been sent to requested attendees, including required and optional attendees, and resources], [MeetingRequestWasSent is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16503");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16503
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.OOF,
                calendar.LegacyFreeBusyStatus,
                16503,
                @"[In t:CalendarItemType Complex Type] The LegacyFreeBusyStatus which value is ""OOF"" specifies the status as Out of Office (OOF).");

            #endregion

            #endregion

            #region Organizer deletes the single calendar item
            ResponseMessageType removedItem = this.DeleteSingleCalendarItem(Role.Organizer, deletedItem, CalendarItemCreateOrDeleteOperationType.SendToNone);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R619");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R610
            this.Site.CaptureRequirementIfIsNotNull(
                removedItem,
                619,
                "[In Messages] DeleteItemSoapIn: For each item being deleted that is not a recurring calendar item, the ItemIds element MUST contain an ItemId child element ([MS-OXWSCORE] section 2.2.4.11).");

            #endregion

            #region Organizer checks whether the calendar item has been really deleted.
            Site.Assert.IsNull(
                this.SearchDeletedSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem.UID),
                "The removed calendar item should not exist in Organizer's calendar folder.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test getting calendar item, meeting request message, meeting response message for accept and meeting cancellation message successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC02_CreateAndAcceptMeeting()
        {
            #region Define a meeting to be created
            int timeInterval = this.TimeInterval;
            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = DateTime.Now.AddHours(timeInterval);

            // Indicates the Start property is serialized in the SOAP message.
            meetingItem.StartSpecified = true;
            timeInterval++;
            meetingItem.End = DateTime.Now.AddHours(timeInterval);
            meetingItem.EndSpecified = true;
            meetingItem.Location = this.Location;
            meetingItem.When = string.Format("{0} to {1}", meetingItem.Start.ToString(), meetingItem.End.ToString());
            meetingItem.IsAllDayEvent = true;
            meetingItem.IsAllDayEventSpecified = true;
            meetingItem.IsResponseRequested = false;
            meetingItem.IsResponseRequestedSpecified = true;
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            meetingItem.IsOnlineMeeting = true;
            meetingItem.IsOnlineMeetingSpecified = true;
            meetingItem.AllowNewTimeProposal = true;
            meetingItem.AllowNewTimeProposalSpecified = true;
            meetingItem.ConferenceType = 1;
            meetingItem.ConferenceTypeSpecified = true;
            meetingItem.MeetingWorkspaceUrl = this.MeetingWorkspace;
            meetingItem.NetShowUrl = this.NetShowLocation;
            meetingItem.LegacyFreeBusyStatus = LegacyFreeBusyType.Tentative;
            meetingItem.LegacyFreeBusyStatusSpecified = true;
            #endregion

            #region Organizer creates a meeting with CalendarItemCreateOrDeleteOperationType value set to SendToAllAndSaveCopy
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy);
            Site.Assert.IsNotNull(item, "The meeting should be created successfully.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Request", meetingItem.UID),
                "The meeting request message should be saved to organizer's Sent Items folder after call CreateItem with CalendarItemCreateOrDeleteOperationType set to SendToAllAndSaveCopy.");

            ItemIdType meetingId = item.Items.Items[0].ItemId;
            #endregion

            #region Attendee gets the meeting request message in the inbox and calendar folders respectively
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "The meeting request message should be found in attendee's Inbox folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendToAllAndSaveCopy.");

            #region Capture Code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R488");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R488
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Schedule.Meeting.Request",
                request.ItemClass,
                488,
                "[In CreateItem Operation] This operation [CreateItem] can be used to create meeting request messages.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28504");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28504
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.Tentative,
                request.IntendedFreeBusyStatus,
                28504,
                @"[In t:MeetingRequestMessageType Complex Type] The IntendedFreeBusyStatus which value is ""Tentative"" specifies the status as tentative.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R35501");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R35501
            this.Site.CaptureRequirementIfAreEqual<int>(
                meetingItem.ConferenceType,
                request.ConferenceType,
                35501,
                @"[In t:MeetingRequestMessageType Complex Type] The value of ""ConferenceType"" is ""1"" describes the type of conferencing is presentation");
            #endregion

            AcceptItemType acceptItem = new AcceptItemType();
            acceptItem.ReferenceItemId = new ItemIdType();
            acceptItem.ReferenceItemId.Id = request.ItemId.Id;

            CalendarItemType calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The meeting should be found in attendee's Calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendToAllAndSaveCopy.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16504");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16504
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.Tentative,
                calendar.LegacyFreeBusyStatus,
                16504,
                @"[In t:CalendarItemType Complex Type] The LegacyFreeBusyStatus which value is ""Tentative"" specifies the status as tentative.");
            #endregion

            #region Attendee calls CreateItem to accept the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            item = this.CreateSingleCalendarItem(Role.Attendee, acceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Accept the meeting request should be successful.");
            #endregion

            #region Organizer gets the meeting response message in the Inbox folder
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The meeting response from Attendee should be existed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R489");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R489
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Schedule.Meeting.Resp.Pos",
                response.ItemClass,
                489,
                "[In CreateItem Operation] This operation [CreateItem] can be used to create meeting response messages.");
            #endregion

            #region Organizer deletes the meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, meetingId, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(deletedItem, "Organizer should delete the calendar successfully.");
            #endregion

            #region Attendee finds the meeting cancellation message from his Inbox folder
            MeetingCancellationMessageType cancelledMessage = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNotNull(cancelledMessage, "Attendee should receive a meeting cancellation message after organizer calls DeleteItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            ItemIdType removeItemId = cancelledMessage.ItemId;
            RemoveItemType removeItem = new RemoveItemType();
            removeItem.ReferenceItemId = removeItemId;

            item = this.CreateSingleCalendarItem(Role.Attendee, removeItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "The canceled message should be removed successfully.");
            #endregion

            #region Organizer checks whether the meeting cancellation message is saved to organizer's Sent Items folder
            MeetingCancellationMessageType cancelledMeeting = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNull(cancelledMeeting, "The meeting cancellation message should not be saved to organizer's Sent Items folder.");
            #endregion

            #region Clean up organizer's inbox, sentitems and deleteditems folders, and attendee's sentitems and deleteditems folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test getting recurring calendar item, meeting request message and meeting response message for decline successfully. 
        /// It also verifies the elements related to recurring/occurrence in CalendarItemType and MeetingRequestMessageType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC03_CreateRecurringMeetingAndDecline()
        {
            #region Organizer creates a recurring meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            #region Define a recurring meeting
            int timeInterval = this.TimeInterval;
            DateTime startTime = DateTime.Now.AddHours(timeInterval);

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = this.PatternInterval;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = this.NumberOfOccurrences;
            range.StartDate = startTime;

            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = startTime;

            // Indicates the Start property is serialized in the SOAP message.
            meetingItem.StartSpecified = true;
            timeInterval++;
            meetingItem.End = startTime.AddHours(timeInterval);
            meetingItem.EndSpecified = true;
            meetingItem.Location = this.Location;
            meetingItem.Recurrence = new RecurrenceType();
            meetingItem.Recurrence.Item = pattern;
            meetingItem.Recurrence.Item1 = range;
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            // Create the recurring meeting
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Server should return success for creating a recurring meeting.");
            ItemIdType meetingId = item.Items.Items[0].ItemId;

            #region Verify FirstOccurrence and LastOccurrence in the CalendarItemType
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, meetingId);
            Site.Assert.IsNotNull(getItem, "The calendar item to be deleted should exist.");

            string actualFirstOccurrenceId = ((CalendarItemType)getItem.Items.Items[0]).FirstOccurrence.ItemId.Id;
            string actualLastOccurrenceId = ((CalendarItemType)getItem.Items.Items[0]).LastOccurrence.ItemId.Id;

            // Get the first occurrence by item id
            OccurrenceItemIdType firstOccurrenceId = new OccurrenceItemIdType();
            firstOccurrenceId.ChangeKey = meetingId.ChangeKey;
            firstOccurrenceId.RecurringMasterId = meetingId.Id;
            firstOccurrenceId.InstanceIndex = 1;

            getItem = this.GetSingleCalendarItem(Role.Organizer, firstOccurrenceId);
            Site.Assert.IsNotNull(getItem, "The first occurrence should be found.");
            CalendarItemType firstOccurrence = getItem.Items.Items[0] as CalendarItemType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R213");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R213
            this.Site.CaptureRequirementIfAreEqual<string>(
                firstOccurrence.ItemId.Id,
                actualFirstOccurrenceId,
                213,
                "[In t:CalendarItemType Complex Type]FirstOccurrence: Specifies the first occurrence of a recurring calendar item.");

            // Get the last occurrence by item id
            OccurrenceItemIdType lastOccurrenceId = new OccurrenceItemIdType();
            lastOccurrenceId.ChangeKey = meetingId.ChangeKey;
            lastOccurrenceId.RecurringMasterId = meetingId.Id;
            lastOccurrenceId.InstanceIndex = this.NumberOfOccurrences;

            getItem = this.GetSingleCalendarItem(Role.Organizer, lastOccurrenceId);
            Site.Assert.IsNotNull(getItem, "The last occurrence should be found.");
            CalendarItemType lastOccurrence = getItem.Items.Items[0] as CalendarItemType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R215");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R215
            this.Site.CaptureRequirementIfAreEqual<string>(
                lastOccurrence.ItemId.Id,
                actualLastOccurrenceId,
                215,
                "[In t:CalendarItemType Complex Type]LastOccurrence: Specifies the last occurrence of a recurring calendar item.");
            #endregion
            #endregion

            #region Organizer deletes an occurrence of the recurring calendar item with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            OccurrenceItemIdType occurrenceItemId = new OccurrenceItemIdType();
            occurrenceItemId.ChangeKey = meetingId.ChangeKey;
            occurrenceItemId.RecurringMasterId = meetingId.Id;
            occurrenceItemId.InstanceIndex = this.InstanceIndex;

            getItem = this.GetSingleCalendarItem(Role.Organizer, occurrenceItemId);
            Site.Assert.IsNotNull(getItem, "The calendar item to be deleted should exist.");

            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, occurrenceItemId, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(deletedItem, "The occurrence of the recurring calendar item should be deleted.");
            #endregion

            #region Attendee calls CreateItem to decline the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "Attendee should receive the meeting request message in the Inbox folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            bool isFirstOccurrence = (firstOccurrence.Start == request.FirstOccurrence.Start) && (firstOccurrence.End == request.FirstOccurrence.End);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R341, expected value of Start of first occurrence is {0} and actual value is {1}; expected value of End of first occurrence is {2} and actual value is {3}", firstOccurrence.Start, request.FirstOccurrence.Start, firstOccurrence.End, request.FirstOccurrence.End);

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R341
            this.Site.CaptureRequirementIfIsTrue(
                isFirstOccurrence,
                341,
                "[In t:MeetingRequestMessageType Complex Type] FirstOccurrence: Represents the first occurrence of a recurring meeting item.");

            bool isLastOccurrence = (lastOccurrence.Start == request.LastOccurrence.Start) && (lastOccurrence.End == request.LastOccurrence.End);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R343, expected value of Start of last occurrence is {0} and actual value is {1}; expected value of End of last occurrence is {2} and actual value is {3}", lastOccurrence.Start, request.LastOccurrence.Start, lastOccurrence.End, request.LastOccurrence.End);

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R343
            this.Site.CaptureRequirementIfIsTrue(
                isLastOccurrence,
                343,
                "[In t:MeetingRequestMessageType Complex Type] LastOccurrence: Represents the last occurrence of a recurring meeting item.");

            DeclineItemType declinedItem = new DeclineItemType();
            declinedItem.ReferenceItemId = new ItemIdType();
            declinedItem.ReferenceItemId.Id = request.ItemId.Id;

            item = this.CreateSingleCalendarItem(Role.Attendee, declinedItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Decline meeting request should be successful.");
            #endregion

            #region Organizer calls CreateItem with MeetingCancellationMessageType to cancel the created meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The decline response should be in the Inbox folder.");
            CancelCalendarItemType cancelCalendarItem = new CancelCalendarItemType();
            cancelCalendarItem.ReferenceItemId = response.AssociatedCalendarItemId;

            item = this.CreateSingleCalendarItem(Role.Organizer, cancelCalendarItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Cancel the meeting through CancelCalendarItemType using CreateItem should be successful.");
            #endregion

            #region Attendee finds the meeting cancellation message from his Inbox folder
            MeetingCancellationMessageType canceledMeeting = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNotNull(canceledMeeting, "The canceled meeting should be in the Inbox folder.");
            #endregion

            #region Clean up organizer's inbox, sentitems and deleteditems folders, and attendee's inbox, sentitems and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        ///  This test case is designed to test getting calendar item, meeting request message, meeting response message for tentatively accepted and meeting cancellation message successfully.
        ///  It also verifies the elements related to adjacent/conflicting meeting in CalendarItemType and MeetingRequestMessageType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC04_CreateMeetingAndTentativelyAccept()
        {
            #region Define a meeting to be created by organizer
            DateTime meetingStart = DateTime.UtcNow.AddMonths(1);
            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = meetingStart;
            meetingItem.StartSpecified = true;
            meetingItem.End = meetingStart.AddHours(1);
            meetingItem.EndSpecified = true;
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            #region Attendee creates calendar items for triggering the adjacent/conflicting meeting
            int timeInterval = 1;
            CalendarItemType adjacentCalendar = new CalendarItemType();
            adjacentCalendar.Subject = Common.GenerateResourceName(this.Site, "AdjacentCalendar");
            adjacentCalendar.Start = meetingStart.AddHours(timeInterval);
            adjacentCalendar.StartSpecified = true;
            timeInterval++;
            adjacentCalendar.End = meetingStart.AddHours(timeInterval);
            adjacentCalendar.EndSpecified = true;

            CalendarItemType conflictCalendar = new CalendarItemType();
            conflictCalendar.Subject = Common.GenerateResourceName(this.Site, "ConflictCalendar");
            conflictCalendar.Start = meetingStart;
            conflictCalendar.StartSpecified = true;
            conflictCalendar.End = meetingStart.AddHours(this.TimeInterval);
            conflictCalendar.EndSpecified = true;

            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Attendee, conflictCalendar, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "The conflict calendar should be created successfully.");
            item = this.CreateSingleCalendarItem(Role.Attendee, adjacentCalendar, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "The adjacent calendar should be created successfully.");
            #endregion

            #region Organizer creates the meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Organizer creates the meeting item should be successful.");
            #endregion

            #region Attendee finds the meeting request message from his Inbox folder
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R630");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R630
            this.Site.CaptureRequirementIfIsNotNull(
                request,
                630,
                @"[In Messages] GetItemSoapIn: For each item being retrieved that is not a recurring calendar item, the ItemIds element MUST contain an ItemId child element ([MS-OXWSCORE] section 2.2.4.11).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R323");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R323
            // Because of Attendee only create one adjacent meeting in this case, therefore, excepted value '1'.
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                request.AdjacentMeetingCount,
                323,
                @"[In t:MeetingRequestMessageType Complex Type] AdjacentMeetingCount: Represents the total number of calendar items that are adjacent to the meeting time.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R327");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R327
            this.Site.CaptureRequirementIfAreEqual<string>(
                adjacentCalendar.Subject,
                request.AdjacentMeetings.Items[0].Subject,
                327,
                @"[In t:MeetingRequestMessageType Complex Type] AdjacentMeetings: Identifies all calendar items that are adjacent to the meeting time.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R321");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R321
            // Because of Attendee only create one conflicting meeting in this case, therefore, excepted value '1'.
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                request.ConflictingMeetingCount,
                321,
                @"[In t:MeetingRequestMessageType Complex Type] ConflictingMeetingCount: Represents the number of calendar items that conflict with the meeting item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R325");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R325
            this.Site.CaptureRequirementIfAreEqual<string>(
                conflictCalendar.Subject,
                request.ConflictingMeetings.Items[0].Subject,
                325,
                @"[In t:MeetingRequestMessageType Complex Type] ConflictingMeetings: Identifies all calendar items that conflict with the meeting time.");
            #endregion

            #region Attendee tentatively accepts the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            TentativelyAcceptItemType tentativelyAcceptItem = new TentativelyAcceptItemType();
            tentativelyAcceptItem.ReferenceItemId = new ItemIdType();
            tentativelyAcceptItem.ReferenceItemId.Id = request.ItemId.Id;
            item = this.CreateSingleCalendarItem(Role.Attendee, tentativelyAcceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Tentatively accept the meeting should be successful.");
            #endregion

            #region Organizer gets the meeting response message from his Inbox folder
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "Organizer should receive the meeting response message after attendee tentatively accept the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R81");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R81
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Tentative,
                response.ResponseType,
                81,
                @"[In t:ResponseTypeType Simple Type] Tentative: Indicates that the recipient has tentatively accepted the meeting.");
            #endregion

            #region Clean up organizer's inbox, calendar and deleteditems folders, and attendee's sentitems, calendar and deleteditems folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemCreateOrDeleteOperationType set to the value of SendOnlyToAll when it used in CreateItem and DeleteItem operation. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC05_CreateAndDeleteCalendarItemWithSendOnlyToAll()
        {
            // Verify CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll when it used in CreateItem and DeleteItem operation.
            this.VerifyCalendarItemCreateOrDeleteOperationType(CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);

            #region Clean up organizer's deleteditems folder, and attendee's inbox, calendar and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemCreateOrDeleteOperationType set to the value of SendToAllAndSaveCopy when it used in CreateItem and DeleteItem operation. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC06_CreateAndDeleteCalendarItemWithSendToAllAndSaveCopy()
        {
            // Verify CalendarItemCreateOrDeleteOperationType set to SendToAllAndSaveCopy when it used in CreateItem and DeleteItem operation.
            this.VerifyCalendarItemCreateOrDeleteOperationType(CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy);

            #region Clean up organizer's sentitems and deleteditems folder, and attendee's inbox, calendar and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test CalendarItemCreateOrDeleteOperationType set to the value of SendToNone when it used in CreateItem and DeleteItem operation. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC07_CreateAndDeleteCalendarItemWithSendToNone()
        {
            // Verify CalendarItemCreateOrDeleteOperationType set to SendToNone when it used in CreateItem and DeleteItem operation.
            this.VerifyCalendarItemCreateOrDeleteOperationType(CalendarItemCreateOrDeleteOperationType.SendToNone);
        }

        /// <summary>
        /// This test case is designed to verify CalendarItemType, MeetingRequestMessageType and MeetingResponseMessageType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC08_CreateAndAcceptSingleMeeting()
        {
            #region Organizer creates a meeting
            #region Set the properties of the meeting to create
            CalendarItemType meeting = new CalendarItemType();
            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meeting.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };

            meeting.Subject = this.Subject;
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Start = DateTime.UtcNow.AddDays(1);
            meeting.StartSpecified = true;
            meeting.End = meeting.Start.AddHours(2);
            meeting.EndSpecified = true;
            meeting.LegacyFreeBusyStatus = LegacyFreeBusyType.Busy;
            meeting.LegacyFreeBusyStatusSpecified = true;
            meeting.Location = this.Location;
            meeting.When = string.Format("{0} to {1}", meeting.Start.ToString(), meeting.End.ToString());
            meeting.IsAllDayEvent = false;
            meeting.IsAllDayEventSpecified = true;
            meeting.IsResponseRequested = true;
            meeting.IsResponseRequestedSpecified = true;
            meeting.IsOnlineMeeting = true;
            meeting.IsOnlineMeetingSpecified = true;
            meeting.ConferenceType = 2;
            meeting.ConferenceTypeSpecified = true;
            meeting.AllowNewTimeProposal = true;
            meeting.AllowNewTimeProposalSpecified = true;
            meeting.MeetingWorkspaceUrl = this.MeetingWorkspace;
            meeting.NetShowUrl = this.NetShowLocation;
            #endregion

            #region Create the meeting and sends it to all attendees
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The meeting should be created successfully.");
            #endregion

            #region Get and verify the CalendarItemType of created meeting
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, item.Items.Items[0].ItemId);
            Site.Assert.IsNotNull(getItem, "The created calendar should exist.");

            CalendarItemType createdCalendarItem = getItem.Items.Items[0] as CalendarItemType;

            #region Verify the child elements of CalendarItemType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R151");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R151
            // The UID of the meeting item was created with System.Guid.NewGuid() method that can guarantee the uniqueness.
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.UID,
                createdCalendarItem.UID,
                151,
                @"[In t:CalendarItemType Complex Type] UID: Contains the unique identifier for the calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R157");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R157
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                meeting.Start.Date,
                createdCalendarItem.Start.Date,
                157,
                @"[In t:CalendarItemType Complex Type] Start: Specifies the start date and time of a duration.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R159");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R159
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                meeting.End.Date,
                createdCalendarItem.End.Date,
                159,
                @"[In t:CalendarItemType Complex Type] End: Specifies the end date and time of a duration.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R721");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R721
            this.Site.CaptureRequirementIfIsFalse(
                createdCalendarItem.IsAllDayEvent,
                721,
                @"[In t:CalendarItemType Complex Type] otherwise [if calendar item or meeting request does not represent an all-day event], [IsAllDayEvent is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16500");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16500
            this.Site.CaptureRequirementIfAreEqual(
                LegacyFreeBusyType.Busy,
                createdCalendarItem.LegacyFreeBusyStatus,
                16500,
                @"[In t:CalendarItemType Complex Type] The LegacyFreeBusyStatus which value is ""Busy"" specifies the status as busy.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R167");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R167
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.Location.ToLower(),
                createdCalendarItem.Location.ToLower(),
                167,
                @"[In t:CalendarItemType Complex Type] Location: Specifies the location of a meeting or appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R169");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R169
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.When,
                createdCalendarItem.When,
                169,
                @"[In t:CalendarItemType Complex Type] When: Provides information about when a calendar or meeting item occurs.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R728");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R728
            this.Site.CaptureRequirementIfIsTrue(
                createdCalendarItem.IsMeeting,
                728,
                @"[In t:CalendarItemType Complex Type] [IsMeeting is] True if the calendar item is a meeting or appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R731");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R731
            this.Site.CaptureRequirementIfIsFalse(
                createdCalendarItem.IsCancelled,
                731,
                @"[In t:CalendarItemType Complex Type] otherwise [if a meeting has not been canceled], [IsCancelled is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R733");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R733
            this.Site.CaptureRequirementIfIsFalse(
                createdCalendarItem.IsRecurring,
                733,
                @"[In t:CalendarItemType Complex Type] otherwise [if a calendar item is not part of a recurring item], [IsRecurring is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R734");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R734
            this.Site.CaptureRequirementIfIsTrue(
                createdCalendarItem.MeetingRequestWasSent,
                734,
                @"[In t:CalendarItemType Complex Type] [MeetingRequestWasSent is] True, if meeting request has been sent to requested attendees, including required and optional attendees, and resources.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R736");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R736
            this.Site.CaptureRequirementIfIsTrue(
                createdCalendarItem.IsResponseRequested,
                736,
                @"[In t:CalendarItemType Complex Type] [IsResponseRequested is] True, if a response to an item is requested.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R512");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R512, "Organizer" specified that the recipient is the meeting organizer in this calendar item currently.
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Organizer,
                createdCalendarItem.MyResponseType,
                512,
                @"[In t:CalendarItemType Complex Type] MyResponseType: Specifies the status of the response to a calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R80");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R80
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Organizer,
                createdCalendarItem.MyResponseType,
                80,
                @"[In t:ResponseTypeType Simple Type] Organizer: Indicates that the recipient is the meeting organizer.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R181");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R181
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.Single,
                createdCalendarItem.CalendarItemType1,
                181,
                @"[In t:CalendarItemType Complex Type] CalendarItemType: Specifies the occurrence type of a calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R185");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R185
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.OrganizerEmailAddress.ToLower(),
                createdCalendarItem.Organizer.Item.EmailAddress.ToLower(),
                185,
                @"[In t:CalendarItemType Complex Type] Organizer: Specifies the organizer of a meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R187");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R187
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.AttendeeEmailAddress.ToLower(),
                createdCalendarItem.RequiredAttendees[0].Mailbox.EmailAddress.ToLower(),
                187,
                @"[In t:CalendarItemType Complex Type] RequiredAttendees: Specifies attendees that are required to attend a meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R189");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R189
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.OrganizerEmailAddress.ToLower(),
                createdCalendarItem.OptionalAttendees[0].Mailbox.EmailAddress.ToLower(),
                189,
                @"[In t:CalendarItemType Complex Type] OptionalAttendees: Specifies attendees who are not required to attend a meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R191");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R191
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.RoomEmailAddress.ToLower(),
                createdCalendarItem.Resources[0].Mailbox.EmailAddress.ToLower(),
                191,
                @"[In t:CalendarItemType Complex Type] Resources: Specifies a scheduled resource for a meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R209");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R209, "1" specified the calendar item on the organizer's calendar represents a meeting
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                createdCalendarItem.AppointmentState,
                209,
                @"[In t:CalendarItemType Complex Type] AppointmentState: Specifies the status of the appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R738");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R738
            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                createdCalendarItem.AppointmentState,
                738,
                @"[In t:CalendarItemType Complex Type] [AppointmentState: Valid values include:] 1: the calendar item on the organizer's calendar represents a meeting");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R227");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R227
            this.Site.CaptureRequirementIfAreEqual<int>(
                meeting.ConferenceType,
                createdCalendarItem.ConferenceType,
                227,
                @"[In t:CalendarItemType Complex Type]ConferenceType: Specifies the type of conferencing that is performed with a calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R743");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R743
            this.Site.CaptureRequirementIfAreEqual<int>(
                2,
                createdCalendarItem.ConferenceType,
                743,
                @"[In t:CalendarItemType Complex Type] [ConferenceType: Valid values include:] 2: chat");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R746");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R746
            this.Site.CaptureRequirementIfIsTrue(
                createdCalendarItem.IsOnlineMeeting,
                746,
                @"[In t:CalendarItemType Complex Type] [IsOnlineMeeting is] True, if the meeting is online.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R233");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R233
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.MeetingWorkspaceUrl.ToLower(),
                createdCalendarItem.MeetingWorkspaceUrl.ToLower(),
                233,
                @"[In t:CalendarItemType Complex Type] MeetingWorkspaceUrl: Contains the URL for the Meeting Workspace that is included in the calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R235");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R235
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.NetShowUrl,
                createdCalendarItem.NetShowUrl,
                235,
                @"[In t:CalendarItemType Complex Type] NetShowUrl: Specifies the URL for an online meeting.");

            if (Common.IsRequirementEnabled(699, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R699");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R699
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.Start.Date,
                    createdCalendarItem.StartWallClock.Date,
                    699,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""StartWallClock"" with type ""xs:dateTime"" which specifies the start time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(700, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R700");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R700
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.End.Date,
                    createdCalendarItem.EndWallClock.Date,
                    700,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""EndWallClock"" with type ""xs:dateTime"" which specifies the ending time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(80048, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R800480");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R800480
                this.Site.CaptureRequirementIfIsTrue(
                    createdCalendarItem.IsOrganizer && createdCalendarItem.IsOrganizerSpecified,
                    800480,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""IsOrganizer"" with type ""xs:boolean"" which is true specifying the current user is the organizer and/or owner of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
            #endregion
            #endregion

            #region Attendee gets and checks the meeting request in the Inbox folder
            MeetingRequestMessageType receivedRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(receivedRequest, "The meeting request should exist in attendee's inbox folder.");

            #region Verify the child elements of MeetingRequestMessageType and MeetingMessageType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R756");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R756
            this.Site.CaptureRequirementIfIsTrue(
                receivedRequest.IsMeeting,
                756,
                @"[In t:MeetingRequestMessageType Complex Type] [IsMeeting is] True, if the calendar item is a meeting or an appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R759");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R759
            this.Site.CaptureRequirementIfIsFalse(
                receivedRequest.IsCancelled,
                759,
                @"[In t:MeetingRequestMessageType Complex Type] otherwise [if the meeting has not been cancelled], [IsCancelled is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R761");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R761
            this.Site.CaptureRequirementIfIsFalse(
                receivedRequest.IsRecurring,
                761,
                @"[In t:MeetingRequestMessageType Complex Type] otherwise [if the meeting is not part of a recurring series of meetings], [IsRecurring is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R762");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R762
            this.Site.CaptureRequirementIfIsTrue(
                receivedRequest.MeetingRequestWasSent,
                762,
                @"[In t:MeetingRequestMessageType Complex Type] [MeetingRequestWasSent is] True, if a meeting request has been sent to requested attendees.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R309");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R309, "Single" Specifies that the item is not associated with a recurring calendar item.
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.Single,
                receivedRequest.CalendarItemType,
                309,
                @"[In t:MeetingRequestMessageType Complex Type] CalendarItemType: Represents the occurrence type of a meeting item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R53");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R53
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.Single,
                receivedRequest.CalendarItemType,
                53,
                @"[In t:CalendarItemTypeType Simple Type] Single: Specifies that the item is not associated with a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R313");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R313
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.OrganizerEmailAddress.ToLower(),
                receivedRequest.Organizer.Item.EmailAddress.ToLower(),
                313,
                @"[In t:MeetingRequestMessageType Complex Type] Organizer: Represents the organizer of the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R337");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R337, "3" specified the organizer's meeting request has been sent; the attendee's meeting request has been received
            this.Site.CaptureRequirementIfAreEqual<int>(
                3,
                receivedRequest.AppointmentState,
                337,
                @"[In t:MeetingRequestMessageType Complex Type] AppointmentState: Specifies the status of the appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R555");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R555
            this.Site.CaptureRequirementIfAreEqual<int>(
                3,
                receivedRequest.AppointmentState,
                555,
                @"[In t:MeetingRequestMessageType Complex Type] [AppointmentState's] Valid values include:
3: the organizer's meeting request has been sent; the attendee's meeting request has been received");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R283");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R283, NewMeetingRequest identifies the meeting request as a new meeting request.
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.NewMeetingRequest,
                receivedRequest.MeetingRequestType,
                283,
                @"[In t:MeetingRequestMessageType Complex Type] MeetingRequestType: Specifies the type of meeting request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R72");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R72
            this.Site.CaptureRequirementIfAreEqual<MeetingRequestTypeType>(
                MeetingRequestTypeType.NewMeetingRequest,
                receivedRequest.MeetingRequestType,
                72,
                @"[In t:MeetingRequestTypeType Simple Type] NewMeetingRequest: Identifies the meeting request as a new meeting request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28500");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28500
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.Busy,
                receivedRequest.IntendedFreeBusyStatus,
                28500,
                @"[In t:MeetingRequestMessageType Complex Type] The IntendedFreeBusyStatus which value is ""Busy"" specifies the status as busy.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R500");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R500
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                meeting.Start.Date,
                receivedRequest.Start.Date,
                500,
                @"[In t:MeetingRequestMessageType Complex Type] Start: Represents the start time of the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R287");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R287
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                meeting.End.Date,
                receivedRequest.End.Date,
                287,
                @"[In t:MeetingRequestMessageType Complex Type] End: Specifies the end of the duration for a single occurrence of a meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R297");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R297
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.Location.ToLower(),
                receivedRequest.Location.ToLower(),
                297,
                @"[In t:MeetingRequestMessageType Complex Type] Location: Represents the location of the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R755");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R755
            this.Site.CaptureRequirementIfIsFalse(
                receivedRequest.IsAllDayEvent,
                755,
                @"[In t:MeetingRequestMessageType Complex Type] otherwise [if the meeting is not an all-day event], [IsAllDayEvent is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R736");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R736
            this.Site.CaptureRequirementIfIsTrue(
                receivedRequest.IsResponseRequested,
                736,
                @"[In t:CalendarItemType Complex Type] [IsResponseRequested is] True, if a response to an item is requested.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R766");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R766
            this.Site.CaptureRequirementIfIsTrue(
                receivedRequest.IsOnlineMeeting,
                766,
                @"[In t:MeetingRequestMessageType Complex Type] [IsOnlineMeeting is] True, if the meeting is online.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R35502");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R35502
            this.Site.CaptureRequirementIfAreEqual<int>(
                meeting.ConferenceType,
                receivedRequest.ConferenceType,
                35502,
                @"[In t:MeetingRequestMessageType Complex Type] The value of ""ConferenceType"" is ""2"" describes the type of conferencing is chat");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R361");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R361
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.MeetingWorkspaceUrl.ToLower(),
                receivedRequest.MeetingWorkspaceUrl.ToLower(),
                361,
                @"[In t:MeetingRequestMessageType Complex Type] MeetingWorkspaceUrl: Contains the URL for the Meeting Workspace that is included in the meeting item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R363");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R363
            this.Site.CaptureRequirementIfAreEqual<string>(
                meeting.NetShowUrl.ToLower(),
                receivedRequest.NetShowUrl.ToLower(),
                363,
                @"[In t:MeetingRequestMessageType Complex Type] NetShowUrl: Specifies the URL for an online meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R315");

            bool isVerifyMeetingRequiredAttendees = string.Equals(receivedRequest.RequiredAttendees[0].Mailbox.EmailAddress, this.OrganizerEmailAddress, StringComparison.OrdinalIgnoreCase)
                && string.Equals(receivedRequest.RequiredAttendees[1].Mailbox.EmailAddress, this.AttendeeEmailAddress, StringComparison.OrdinalIgnoreCase)
                && receivedRequest.RequiredAttendees.Length == 2;

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R315
            this.Site.CaptureRequirementIfIsTrue(
                isVerifyMeetingRequiredAttendees,
                315,
                @"[In t:MeetingRequestMessageType Complex Type] RequiredAttendees: Represents attendees that are required to attend the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R317");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R317
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.OrganizerEmailAddress.ToLower(),
                receivedRequest.OptionalAttendees[0].Mailbox.EmailAddress.ToLower(),
                317,
                @"[In t:MeetingRequestMessageType Complex Type] OptionalAttendees: Represents attendees who are not required to attend the meeting.");

            if (Common.IsRequirementEnabled(708, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R708");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R708
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.Start.Date,
                    receivedRequest.StartWallClock.Date,
                    708,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""StartWallClock""with type ""xs:dateTime"" which specifies the start time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(709, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R709");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R709
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.End.Date,
                    receivedRequest.EndWallClock.Date,
                    709,
                    @"[In Appendix C: Product Behavior] Implementation does support the complex type ""EndWallClock"" with type ""xs:dateTime"" which specifies the ending time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            AcceptItemType acceptItem = new AcceptItemType();
            acceptItem.ReferenceItemId = new ItemIdType();
            acceptItem.ReferenceItemId.Id = receivedRequest.ItemId.Id;
            #endregion

            #region Attendee accepts the meeting request
            item = this.CreateSingleCalendarItem(Role.Attendee, acceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The response to the meeting request should be successful.");

            CalendarItemType calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The accepted calendar should be found in attendee's calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R205");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R205
            this.Site.CaptureRequirementIfIsNotNull(
                calendar.AppointmentReplyTime,
                205,
                @"[In t:CalendarItemType Complex Type] AppointmentReplyTime: Specifies the date and time that an attendee replied to a meeting request.");
            #endregion

            #region Organizer gets and checks the meeting response from attendeeType
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meeting.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The meeting response message from attendee should exist in organizer's inbox folder.");

            #region Verify the child elements of MeetingResponseMessageType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R82");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R82
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Accept,
                response.ResponseType,
                82,
                @"[In t:ResponseTypeType Simple Type] Accept: Indicates that the recipient accepted the meeting.");

            if (Common.IsRequirementEnabled(906, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R906");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R906
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.Start.Date,
                    response.Start.Date,
                    906,
                    @"[In Appendix C: Product Behavior] Implementation does support Start which is a dateTime element that represents the start time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(907, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R907");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R907
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.End.Date,
                    response.End.Date,
                    907,
                    @"[In Appendix C: Product Behavior] Implementation does support End which is a dateTime element that represents the ending time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(908, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R908");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R908
                this.Site.CaptureRequirementIfAreEqual<string>(
                    meeting.Location.ToLower(),
                    response.Location.ToLower(),
                    908,
                    @"[In Appendix C: Product Behavior] Implementation does support Location which is a string element that represents the location for the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(910, this.Site))
            {
                CalendarItemTypeType actual;
                Site.Assert.IsTrue(Enum.TryParse<CalendarItemTypeType>(response.CalendarItemType, out actual), "The current value of CalendarItemType property should be one of CalendarItemTypeType enum values.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R910");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R910
                this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                    CalendarItemTypeType.Single,
                    actual,
                    910,
                    @"[In Appendix C: Product Behavior] Implementation does support CalendarItemType which is a string element that represents the type of calendar item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The accepted calendar should be found in organizer's calendar folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R141");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R141
            this.Site.CaptureRequirementIfIsNotNull(
                calendar.RequiredAttendees[0].LastResponseTime,
                141,
                @"[In t:AttendeeType Complex Type]LastResponseTime: Specifies the date and time that the latest meeting invitation response was received by the meeting organizer from the meeting attendee.");
            #endregion

            #region Clean up organizer's inbox, calendar and deleteditems folders, and attendee's sentitems, calendar and deleteditems folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify MeetingCancellationMessageType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC09_CreateAndCancelMeeting()
        {
            #region Organizer creates a meeting
            #region Set the properties of the meeting to create
            CalendarItemType meeting = new CalendarItemType();
            meeting.Subject = this.Subject;
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Start = DateTime.UtcNow.AddDays(1);
            meeting.StartSpecified = true;
            meeting.End = meeting.Start.AddHours(2);
            meeting.EndSpecified = true;
            meeting.Location = this.Location;
            meeting.ConferenceType = 0;
            meeting.ConferenceTypeSpecified = true;
            meeting.AllowNewTimeProposal = false;
            meeting.AllowNewTimeProposalSpecified = true;
            meeting.IsOnlineMeeting = false;
            meeting.IsOnlineMeetingSpecified = true;
            meeting.LegacyFreeBusyStatus = LegacyFreeBusyType.Free;
            meeting.LegacyFreeBusyStatusSpecified = true;

            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meeting.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            #region Create the meeting and sends it to all attendees
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Create a meeting item should be successful.");
            #endregion
            #endregion

            #region Attendee gets and declines the meeting request in the Inbox folder
            CalendarItemType calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar item should be found in attendee's Calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R16501");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R16501
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.Free,
                calendar.LegacyFreeBusyStatus,
                16501,
                @"[In t:CalendarItemType Complex Type] The LegacyFreeBusyStatus which value is ""Free"" specifies the status as free.");

            #region Decline the meeting request
            MeetingRequestMessageType receivedRequest = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(receivedRequest, "The meeting request message should exist in attendee's inbox folder.");

            #region Capture Code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R35500");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R35500
            this.Site.CaptureRequirementIfAreEqual<int>(
                meeting.ConferenceType,
                receivedRequest.ConferenceType,
                35500,
                @"[In t:MeetingRequestMessageType Complex Type] The value of ""ConferenceType"" is ""0"" (zero) describes the type of conferencing is video conference");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R767");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R767
            this.Site.CaptureRequirementIfIsFalse(
                receivedRequest.IsOnlineMeeting,
                767,
                @"[In t:MeetingRequestMessageType Complex Type] otherwise [if the meeting is not online], [IsOnlineMeeting is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28501");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28501
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.Free,
                receivedRequest.IntendedFreeBusyStatus,
                28501,
                @"[In t:MeetingRequestMessageType Complex Type] The IntendedFreeBusyStatus which value is ""Free"" specifies the status as free.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R329");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R329
            // The format of the value of Duration elements followed xs:duration (as specified in [XMLSCHEMA2]), because of the duration set to 2 hours when create the meeting, therefore the expected value is "PT2H".
            this.Site.CaptureRequirementIfAreEqual<string>(
                "PT2H",
                receivedRequest.Duration,
                329,
                @"[In t:MeetingRequestMessageType Complex Type] Duration: Represents the duration of the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R331");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R331
            this.Site.CaptureRequirementIfIsNotNull(
                receivedRequest.TimeZone,
                331,
                @"[In t:MeetingRequestMessageType Complex Type] TimeZone: Provides a text description of a time zone.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R335");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R335
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                receivedRequest.AppointmentSequenceNumber,
                335,
                @"[In t:MeetingRequestMessageType Complex Type] AppointmentSequenceNumber: Specifies the sequence number of a version of an appointment.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R137");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R137
            Site.CaptureRequirementIfIsNotNull(
                receivedRequest,
                137,
                @"[In t:AttendeeType Complex Type]Mailbox:  Specifies a fully resolved e-mail address.");
            #endregion

            DeclineItemType declineItem = new DeclineItemType();
            declineItem.ReferenceItemId = new ItemIdType();
            declineItem.ReferenceItemId = receivedRequest.ItemId;

            item = this.CreateSingleCalendarItem(Role.Attendee, declineItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Decline meeting request should be successful.");
            #endregion
            #endregion

            #region Organizer gets the meeting response message and verify ResponseTypeType set to decline
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meeting.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "Organizer should receive the meeting response message after attendee declines the meeting.");

            #region Capture Code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R83");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R83
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Decline,
                response.ResponseType,
                83,
                @"[In t:ResponseTypeType Simple Type] Decline: Indicates that the recipient declined the meeting.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R139");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R139
            this.Site.CaptureRequirementIfAreEqual<ResponseTypeType>(
                ResponseTypeType.Decline,
                response.ResponseType,
                139,
                @"[In t:AttendeeType Complex Type]ResponseType: Specifies the meeting invitation response received for by the meeting organizer from a meeting attendee.");
            #endregion

            CancelCalendarItemType cancelMeetingItem = new CancelCalendarItemType();
            cancelMeetingItem.ReferenceItemId = response.AssociatedCalendarItemId;
            #endregion

            #region Organizer cancels the meeting
            item = this.CreateSingleCalendarItem(Role.Organizer, cancelMeetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R491");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R491
            Site.CaptureRequirementIfIsNotNull(
                item,
                491,
                @"[In CreateItem Operation] It [CreateItem operation] can also be used to cancel a meeting.");
            #endregion

            #region Attendee removes the canceled meeting
            MeetingCancellationMessageType meetingResponse = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meeting.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNotNull(meetingResponse, "Attendee should receive the meeting cancellation message after organizer calls CreateItem to create MeetingCancellationMessage with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R490");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R490
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Schedule.Meeting.Canceled",
                meetingResponse.ItemClass,
                490,
                @"[In CreateItem Operation] This operation [CreateItem] can be used to create meeting cancellation messages.");

            #region Verify the child elements of MeetingResponseMessageType
            if (Common.IsRequirementEnabled(900, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R900");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R900
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.Start.Date,
                    meetingResponse.Start.Date,
                    900,
                    @"[In Appendix C: Product Behavior] Implementation does support Start which is a dateTime element that represents the start time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(901, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R901");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R901
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    meeting.End.Date,
                    meetingResponse.End.Date,
                    901,
                    @"[In Appendix C: Product Behavior] Implementation does support End which is a dateTime element that represents the ending time of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(902, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R902");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R902
                this.Site.CaptureRequirementIfAreEqual<string>(
                    meeting.Location.ToLower(),
                    meetingResponse.Location.ToLower(),
                    902,
                    @"[In Appendix C: Product Behavior] Implementation does support Location which is a string element that represents the location of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(904, this.Site))
            {
                CalendarItemTypeType actual;
                Site.Assert.IsTrue(Enum.TryParse<CalendarItemTypeType>(meetingResponse.CalendarItemType, out actual), "The current value of CalendarItemType should be able to convert into one of CalendarItemTypeType enum values.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R904");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R904
                this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                    CalendarItemTypeType.Single,
                    actual,
                    904,
                    @"[In Appendix C: Product Behavior] Implementation does support CalendarItemType which is a string element that represents the type of calendar item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            RemoveItemType removeItem = removeItem = new RemoveItemType();
            removeItem.ReferenceItemId = new ItemIdType();
            removeItem.ReferenceItemId = meetingResponse.ItemId;

            #region Remove the canceled meeting
            item = this.CreateSingleCalendarItem(Role.Attendee, removeItem, CalendarItemCreateOrDeleteOperationType.SendToNone);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R492");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R492
            this.Site.CaptureRequirementIfIsNotNull(
                item,
                492,
                @"[In CreateItem Operation] and when a meeting is cancelled, it [CreateItem Operation] can be used to remove the meeting item and corresponding meeting cancellation message from the server.");
            #endregion
            #endregion

            #region Clean up organizer's inbox, sentitems and deleteditems folders, and attendee's inbox, sentitems and deleteditems folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as DailyRecurrencePatternType and NumberedRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC10_ModifyOccurrenceWithDailyPatternAndNumberedRange()
        {
            // Verify DailyRecurrencePatternType and NumberedRecurrenceRangeType.
            DailyRecurrencePatternType dailyPattern = new DailyRecurrencePatternType();
            NumberedRecurrenceRangeType numberedRange = new NumberedRecurrenceRangeType();
            numberedRange.NumberOfOccurrences = this.NumberOfOccurrences;
            this.VerifyModifiedOccurrences(dailyPattern, numberedRange);
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as WeeklyRecurrencePatternType and EndDateRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC11_ModifyOccurrenceWithWeeklyPatternAndEndDateRange()
        {
            // Verify WeeklyRecurrencePatternType and EndDateRecurrenceRangeType.
            WeeklyRecurrencePatternType weeklyPattern = new WeeklyRecurrencePatternType();
            weeklyPattern.DaysOfWeek = "Tuesday";
            EndDateRecurrenceRangeType endDateRange = new EndDateRecurrenceRangeType();
            this.VerifyModifiedOccurrences(weeklyPattern, endDateRange);
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as AbsoluteMonthlyRecurrencePatternType and NoEndRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC12_ModifyOccurrenceWithAbsoluteMonthlyPatternAndNoEndRange()
        {
            // Verify AbsoluteMonthlyRecurrencePatternType and NoEndRecurrenceRangeType.
            AbsoluteMonthlyRecurrencePatternType absoluteMonthly = new AbsoluteMonthlyRecurrencePatternType();
            absoluteMonthly.DayOfMonth = 5;
            NoEndRecurrenceRangeType nonEndRange = new NoEndRecurrenceRangeType();
            this.VerifyModifiedOccurrences(absoluteMonthly, nonEndRange);
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as RelativeMonthlyRecurrencePatternType and NumberedRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC13_ModifyOccurrenceWithRelativeMonthlyPatternAndNumberedRange()
        {
            // Verify RelativeMonthlyRecurrencePatternType and NumberedRecurrenceRangeType.
            RelativeMonthlyRecurrencePatternType relativeMonthly = new RelativeMonthlyRecurrencePatternType();
            relativeMonthly.DaysOfWeek = DayOfWeekType.Thursday;
            relativeMonthly.DayOfWeekIndex = DayOfWeekIndexType.First;
            NumberedRecurrenceRangeType numberedRange = new NumberedRecurrenceRangeType();
            numberedRange.NumberOfOccurrences = this.NumberOfOccurrences;
            this.VerifyModifiedOccurrences(relativeMonthly, numberedRange);
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as AbsoluteYearlyRecurrencePatternType and EndDateRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC14_ModifyOccurrenceWithAbsoluteYearlyPatternAndEndDateRange()
        {
            // Verify AbsoluteYearlyRecurrencePatternType and EndDateRecurrenceRangeType.
            AbsoluteYearlyRecurrencePatternType absoluteYearly = new AbsoluteYearlyRecurrencePatternType();
            absoluteYearly.DayOfMonth = 5;
            absoluteYearly.Month = MonthNamesType.February;
            EndDateRecurrenceRangeType endDateRange = new EndDateRecurrenceRangeType();
            this.VerifyModifiedOccurrences(absoluteYearly, endDateRange);
        }

        /// <summary>
        /// This test case is designed to verify an occurrence of a recurring meeting defined as RelativeYearlyRecurrencePatternType and NoEndRecurrenceRangeType.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC15_ModifyOccurrenceWithRelativeYearlyPatternAndNoEndRange()
        {
            // Verify RelativeYearlyRecurrencePatternType and NoEndRecurrenceRangeType.
            RelativeYearlyRecurrencePatternType relativeYearly = new RelativeYearlyRecurrencePatternType();
            relativeYearly.DaysOfWeek = "Wednesday";
            relativeYearly.DayOfWeekIndex = DayOfWeekIndexType.First;
            relativeYearly.Month = MonthNamesType.January;
            NoEndRecurrenceRangeType nonEndRange = new NoEndRecurrenceRangeType();
            this.VerifyModifiedOccurrences(relativeYearly, nonEndRange);
        }

        /// <summary>
        /// This test case is designed to verify properties related to recurring meeting: DeletedOccurrences.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S01_TC16_DeleteOccurrenceOfRecurringMeeting()
        {
            #region Organizer creates a recurring meeting
            // Verify DailyRecurrencePatternType and NumberedRecurrenceRangeType.
            DailyRecurrencePatternType dailyPattern = new DailyRecurrencePatternType();
            NumberedRecurrenceRangeType numberedRange = new NumberedRecurrenceRangeType();
            numberedRange.NumberOfOccurrences = this.NumberOfOccurrences;

            // Define a recurring meeting.
            CalendarItemType meetingItem = this.DefineRecurringMeeting(dailyPattern, numberedRange);
            Site.Assert.IsNotNull(meetingItem, "The meeting item should be created.");

            // Create the recurring meeting.
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The recurring meeting should be created successfully.");
            #endregion

            #region Attendee gets and verifies the recurring meeting request
            CalendarItemType calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The meeting should exist in the attendee's Calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");
            #endregion

            #region Organizer deletes one of the occurrences of the recurring meeting
            // Get the occurrence to be deleted.
            ItemType occurrence = this.GetFirstOccurrenceItem(meetingItem, Role.Organizer);
            Site.Assert.IsNotNull(occurrence, "The specified occurrence item should be found.");

            // Store the start and end time of the occurrence to be deleted.
            CalendarItemType occurrenceItem = occurrence as CalendarItemType;

            // Delete the occurrence.
            bool isDeleted = this.DeleteOccurrenceItem(occurrence.ItemId);
            Site.Assert.IsTrue(isDeleted, "The occurrence item should be deleted.");

            calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar item should exist.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R219");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R219
            this.Site.CaptureRequirementIfIsNotNull(
                calendar.DeletedOccurrences,
                219,
                @"[In t:CalendarItemType Complex Type]DeletedOccurrences: Specifies deleted occurrences of a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R373");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R373
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                occurrenceItem.Start,
                calendar.DeletedOccurrences[0].Start,
                373,
                @"[In t:NonEmptyArrayOfDeletedOccurrencesType Complex Type] DeletedOccurrence: Represents a deleted occurrence of a recurring calendar item.");
            #endregion

            #region Organizer deletes the recurring meeting
            CalendarItemType recurringCalendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(recurringCalendar, "The meeting should exist in the organizer's calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, recurringCalendar.ItemId, CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R620");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R620
            this.Site.CaptureRequirementIfIsNotNull(
                deletedItem,
                620,
                @"[In Messages] DeleteItemSoapIn: For each item being deleted that is a recurring calendar item, the ItemIds element can contain a RecurringMasterItemId child element ([MS-OXWSCORE] section 2.2.4.11) or an OccurrenceItemId child element ([MS-OXWSCORE] section 2.2.4.11).");
            #endregion

            #region Clean up organizer's deleteditems and sentitems folder, and attendee's inbox, calendar and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Delete an occurrence of a recurring meeting.
        /// </summary>
        /// <param name="occurrenceId">The Id of the occurrence to be deleted.</param>
        /// <returns>If delete operation succeeds, return true; otherwise, false.</returns>
        private bool DeleteOccurrenceItem(ItemIdType occurrenceId)
        {
            if (occurrenceId != null)
            {
                ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, occurrenceId, CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy);
                if (deletedItem != null)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Verify ModifiedOccurrences property of a recurring meeting.
        /// </summary>
        /// <param name="pattern">The recurring pattern.</param>
        /// <param name="range">The recurring range.</param>
        private void VerifyModifiedOccurrences(RecurrencePatternBaseType pattern, RecurrenceRangeBaseType range)
        {
            #region Step1: Organizer creates a recurring meeting
            // Define a recurring meeting.
            CalendarItemType meetingItem = this.DefineRecurringMeeting(pattern, range);
            Site.Assert.IsNotNull(meetingItem, "The meeting item should be created first.");

            // Create the recurring meeting.
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "The recurring meeting should be created successfully.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R494");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R494
            this.Site.CaptureRequirementIfIsNotNull(
                item,
                494,
                @"[In CreateItem Operation] This operation [CreateItem] can be used to create meetings.");
            #endregion

            #region Step2: Attendee gets and verifies the recurring meeting request
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int count = 1;

            MeetingRequestMessageType request = null;

            while (request == null && count++ <= upperBound)
            {
               request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
               System.Threading.Thread.Sleep(waitTime);
            }
             
            Site.Assert.IsNotNull(request, "Attendee should receive the meeting request message in the Inbox folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendOnlyToAll.");

            if (Common.IsRequirementEnabled(80048, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R800481");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R800481
                this.Site.CaptureRequirementIfIsFalse(
                    request.IsOrganizer,
                    800481,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""IsOrganizer"" with type ""xs:boolean"" which is false specifying the current user is not the organizer and/or owner of the calendar item. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(903, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R903");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R903
                this.Site.CaptureRequirementIfIsNotNull(
                    request.Recurrence,
                    903,
                    @"[In Appendix C: Product Behavior] Implementation does support Recurrence which is a RecurrenceType element that represents the recurrence of the calendar item. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R28503");
            }

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R28503
            this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                LegacyFreeBusyType.OOF,
                request.IntendedFreeBusyStatus,
                28503,
                @"[In t:MeetingRequestMessageType Complex Type] The IntendedFreeBusyStatus which value is ""OOF"" specifies the status as Out of Office (OOF).");

            if (Common.IsRequirementEnabled(807, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R807");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R807
                this.Site.CaptureRequirementIfIsNotNull(
                    request,
                    807,
                    @"[In Appendix C: Product Behavior] GetItemSoapIn: For each item being retrieved that is a recurring calendar item, implementation does contain a RecurringMasterItemId child element ([MS-OXWSCORE] section 2.2.4.11) or an OccurrenceItemId child element ([MS-OXWSCORE] section 2.2.4.11). (Exchange 2007 and above follow this behavior.)");
            }

            // Verify the calendar item
            CalendarItemType calendar = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar item to be verified should exist in Attendee's Calendar folder.");

            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Attendee, calendar.FirstOccurrence.ItemId);
            CalendarItemType firstOccurrence = getItem.Items.Items[0] as CalendarItemType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R54");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R54
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.Occurrence,
                firstOccurrence.CalendarItemType1,
                54,
                @"[In t:CalendarItemTypeType Simple Type] Occurrence: Specifies that the item is an occurrence of a recurring calendar item.");

            #region Capture Code for CalendarItemType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R742");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R742
            this.Site.CaptureRequirementIfAreEqual<int>(
                meetingItem.ConferenceType,
                calendar.ConferenceType,
                742,
                @"[In t:CalendarItemType Complex Type] [ConferenceType: Valid values include:] 1: presentation");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R744");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R744
            this.Site.CaptureRequirementIfIsTrue(
                calendar.AllowNewTimeProposalSpecified && calendar.AllowNewTimeProposal,
                744,
                @"[In t:CalendarItemType Complex Type] [AllowNewTimeProposal is] True, if a new meeting time can be proposed for a meeting by an attendee.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R56");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R56
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.RecurringMaster,
                calendar.CalendarItemType1,
                56,
                @"[In t:CalendarItemTypeType Simple Type] RecurringMaster: Specifies that the item is the master item that contains the recurrence pattern for a calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R732");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R732
            bool isChecked = calendar.IsRecurringSpecified && calendar.IsRecurring;
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                732,
                @"[In t:CalendarItemType Complex Type] [IsRecurring is] True, if a calendar item is part of a recurring item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R737");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R737
            isChecked = calendar.IsResponseRequestedSpecified && calendar.IsResponseRequested;

            this.Site.CaptureRequirementIfIsFalse(
                isChecked,
                737,
                @"[In t:CalendarItemType Complex Type] otherwise [if a response to an item is not requested], [IsResponseRequested is] false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R739");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R739
            this.Site.CaptureRequirementIfAreEqual<int>(
                3,
                calendar.AppointmentState,
                739,
                @"[In t:CalendarItemType Complex Type] [AppointmentState: Valid values include:] 3: the meeting request corresponding to the calendar item has been received");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R760");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R760
            isChecked = calendar.IsRecurringSpecified && calendar.IsRecurring;
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                760,
                @"[In t:MeetingRequestMessageType Complex Type] [IsRecurring is] True, if the meeting is part of a recurring series of meetings.");

            if (Common.IsRequirementEnabled(701, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R701");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R701
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "Greenwich Standard Time",
                    calendar.StartTimeZoneId,
                    701,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""StartTimeZoneId"" with type ""xs:string"" which specifies the calendar item start time zone identifier. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(702, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R702");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R702
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "Greenwich Standard Time",
                    calendar.EndTimeZoneId,
                    702,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""EndTimeZoneId"" with type ""xs:string"" which specifies the calendar item end time zone identifier. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(703, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R703");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R703
                this.Site.CaptureRequirementIfAreEqual<LegacyFreeBusyType>(
                    LegacyFreeBusyType.OOF,
                    calendar.IntendedFreeBusyStatus,
                    703,
                    @"[In Appendix C: Product Behavior] Implementation does support complex type ""IntendedFreeBusyStatus"" with type ""LegacyFreeBusyType ([MS-OXWSCDATA] section 2.2.3.16)"" which indicates how the organizer of the meeting wants it to show up in the attendee's calendar when the meeting is accepted. (Exchange 2013 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R18407");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R18407
            isChecked = !string.IsNullOrEmpty(calendar.Organizer.Item.EmailAddress);
            this.Site.CaptureRequirementIfIsTrue(
                isChecked,
                18407,
                @"[In t:CalendarItemType Complex Type] When the Mailbox element of Organizer element include an EmailAddress element of t:NonEmptyStringType, the t:NonEmptyStringType simple type specifies a string that MUST have a minimum of one character.");
            #endregion

            // Verify Recurrence
            Site.Assert.IsNotNull(calendar.Recurrence, "The Recurrence property of the calendar should not be null.");
            Site.Assert.IsNotNull(calendar.Recurrence.Item, "The pattern of the calendar should not be null.");
            Site.Assert.IsNotNull(calendar.Recurrence.Item1, "The range of the calendar should not be null.");
            this.VerifyReccurrenceType(calendar.Recurrence);
            #endregion

            #region Step3: Organizer updates one of the occurrences of the recurring meeting
            // Get the occurrence to be updated.
            ItemType occurrence = this.GetFirstOccurrenceItem(meetingItem, Role.Organizer);
            Site.Assert.IsNotNull(occurrence, "The specified occurrence item should be found.");

            // Update the occurrence.
            bool isUpdated = this.UpdateOccurrenceItem(occurrence);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R651");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R651
            this.Site.CaptureRequirementIfIsTrue(
                isUpdated,
                651,
                @"[In Messages] UpdateItemSoapIn: For each item being updated that is a recurring calendar item, the ItemChange element can contain a RecurringMasterItemId child element ([MS-OXWSCORE] section 3.1.4.9.3.7) or an OccurrenceItemId child element ([MS-OXWSCORE] section 3.1.4.9.3.7).");

            CalendarItemType calendarItem = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendarItem, "The calendar item should exist.");
            Site.Assert.IsTrue(calendarItem.CalendarItemType1 == CalendarItemTypeType.RecurringMaster, "The type of the calendar should be RecurringMaster.");
            Site.Assert.IsNotNull(calendarItem.ModifiedOccurrences, "The ModifiedOccurrences should contain one occurrence at least.");

            #region Capture Code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R377");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R377
            this.Site.CaptureRequirementIfAreEqual<string>(
                calendarItem.FirstOccurrence.ItemId.Id,
                calendarItem.ModifiedOccurrences[0].ItemId.Id,
                377,
                @"[In t:NonEmptyArrayOfOccurrenceInfoType Complex Type] Occurrence: Represents a modified occurrence of a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R217");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R217
            this.Site.CaptureRequirementIfAreEqual<string>(
                calendarItem.FirstOccurrence.ItemId.Id,
                calendarItem.ModifiedOccurrences[0].ItemId.Id,
                217,
                @"[In t:CalendarItemType Complex Type]ModifiedOccurrences: Specifies recurring calendar item occurrences that have been modified so that they differ from original occurrences (or instances of the recurring master item).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R381");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R381
            this.Site.CaptureRequirementIfAreEqual<string>(
                calendarItem.FirstOccurrence.ItemId.Id,
                calendarItem.ModifiedOccurrences[0].ItemId.Id,
                381,
                @"[In t:OccurrenceInfoType Complex Type] ItemId: Contains the identifier of a modified occurrence of a recurring calendar item.");

            CalendarItemType occurrenceCalendar = occurrence as CalendarItemType;
            Site.Assert.IsNotNull(occurrenceCalendar, "The type conversion from ItemType to CalendarItemType should succeed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R383");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R383
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                occurrenceCalendar.Start.AddHours(-26.0),
                calendarItem.ModifiedOccurrences[0].Start,
                383,
                @"[In t:OccurrenceInfoType Complex Type] Start: Contains the start time of a modified occurrence of a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R385");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R385
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                occurrenceCalendar.End.AddHours(-26.0),
                calendarItem.ModifiedOccurrences[0].End,
                385,
                @"[In t:OccurrenceInfoType Complex Type] End: Contains the end time of a modified occurrence of a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R387");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R387
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                occurrenceCalendar.OriginalStart,
                calendarItem.ModifiedOccurrences[0].OriginalStart,
                387,
                @"[In t:OccurrenceInfoType Complex Type] OriginalStart: Contains the original start time of a modified occurrence of a recurring calendar item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R161");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R161
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                occurrenceCalendar.Start,
                calendarItem.FirstOccurrence.OriginalStart,
                161,
                @"[In t:CalendarItemType Complex Type] OriginalStart: Represents the original start time of a calendar item (only for occurrences/exceptions).");
            #endregion
            #endregion

            #region Step4: Attendee gets and verifies the modified occurrence
            CalendarItemType updatedOccurrence = null;
            int counter = 0;
            while (counter < this.UpperBound)
            {
                System.Threading.Thread.Sleep(this.WaitTime);
                updatedOccurrence = this.GetFirstOccurrenceItem(meetingItem, Role.Attendee) as CalendarItemType;

                if (updatedOccurrence.Location.ToLower() == this.LocationUpdate.ToLower())
                {
                    break;
                }

                counter++;
            }

            if (counter == this.UpperBound && updatedOccurrence.Location.ToLower() != this.LocationUpdate.ToLower())
            {
                Site.Assert.Fail("Attendee should get the updates after organizer updates the meeting.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R55");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R55
            this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                CalendarItemTypeType.Exception,
                updatedOccurrence.CalendarItemType1,
                55,
                @"[In t:CalendarItemTypeType Simple Type] Exception: Specifies that the item is an exception to a recurring calendar item.");

            Site.Assert.AreEqual<string>(
                this.LocationUpdate,
                updatedOccurrence.Location,
                string.Format("The value of Location property of the updated occurrence is Expected: {0}; Actual: {1}", this.LocationUpdate, updatedOccurrence.Location));
            #endregion

            #region Step5: Clean up organizer's calendar and deleteditems folders, and attendee's inbox, calendar and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// Verify Recurrence property of a recurring meeting
        /// </summary>
        /// <param name="recurrence">An instance of RecurrenceType</param>
        private void VerifyReccurrenceType(RecurrenceType recurrence)
        {
            this.VerifyRecurrencePatternType(recurrence.Item);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R391");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R391
            this.Site.CaptureRequirementIfIsNotNull(
                recurrence.Item1,
                391,
                @"[In t:RecurrenceType Complex Type] The RecurrenceRangeTypes group specifies the recurrence patterns with numbered recurrences, non-ending recurrence patterns, and recurrence patterns with a set start and end date, as specified in [MS-OXWSCDATA] section 2.2.7.2.");
        }

        /// <summary>
        /// Verify recurrence pattern
        /// </summary>
        /// <param name="pattern">An instance of RecurrencePatternBaseType.</param>
        private void VerifyRecurrencePatternType(RecurrencePatternBaseType pattern)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R339");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R339
            this.Site.CaptureRequirementIfIsNotNull(
                pattern,
                339,
                @"[In t:MeetingRequestMessageType Complex Type] Recurrence: Contains the recurrence pattern for meeting items and meeting requests.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R390");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R390
            this.Site.CaptureRequirementIfIsNotNull(
                pattern,
                390,
                @"[In t:RecurrenceType Complex Type] The RecurrencePatternTypes group specifies the recurrence pattern for calendar items and meeting requests, as specified in [MS-OXWSCDATA] section 2.2.7.1.");

            AbsoluteMonthlyRecurrencePatternType absoluteMonthly = pattern as AbsoluteMonthlyRecurrencePatternType;
            if (absoluteMonthly != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R997");

                // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R997
                this.Site.CaptureRequirementIfAreEqual<int>(
                    5,
                    absoluteMonthly.DayOfMonth,
                    "MS-OXWSCDATA",
                    997,
                    @"[In t:AbsoluteMonthlyRecurrencePatternType Complex Type] This property [DayOfMonth] MUST be present.");
            }
            else
            {
                AbsoluteYearlyRecurrencePatternType absoluteYearly = pattern as AbsoluteYearlyRecurrencePatternType;
                if (absoluteYearly != null)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1003");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R1003
                    this.Site.CaptureRequirementIfAreEqual<MonthNamesType>(
                        MonthNamesType.February,
                        absoluteYearly.Month,
                        "MS-OXWSCDATA",
                        1003,
                        @"[In t:AbsoluteYearlyRecurrencePatternType Complex Type] This property [Month] MUST be present.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1001");

                    // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R1001
                    this.Site.CaptureRequirementIfAreEqual<int>(
                        5,
                        absoluteYearly.DayOfMonth,
                        "MS-OXWSCDATA",
                        1001,
                        @"[In t:AbsoluteYearlyRecurrencePatternType Complex Type] This property [DayOfMonth] MUST be present.");
                }
                else
                {
                    RelativeYearlyRecurrencePatternType relativeYearly = pattern as RelativeYearlyRecurrencePatternType;
                    if (relativeYearly != null)
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1263");

                        // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R1263
                        this.Site.CaptureRequirementIfAreEqual<DayOfWeekIndexType>(
                            DayOfWeekIndexType.First,
                            relativeYearly.DayOfWeekIndex,
                            "MS-OXWSCDATA",
                            1263,
                            @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [DayOfWeekIndex] MUST be present.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1261");

                        // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R1261
                        this.Site.CaptureRequirementIfAreEqual<string>(
                            "Wednesday",
                            relativeYearly.DaysOfWeek,
                            "MS-OXWSCDATA",
                            1261,
                            @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [DaysOfWeek] MUST be present.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1265");

                        // Verify MS-OXWSMTGS requirement: MS-OXWSCDATA_R1265
                        this.Site.CaptureRequirementIfAreEqual<MonthNamesType>(
                            MonthNamesType.January,
                            relativeYearly.Month,
                            "MS-OXWSCDATA",
                            1265,
                            @"[In t:RelativeYearlyRecurrencePatternType Complex Type] This element [Month] MUST be present.");
                    }
                }
            }
        }

        /// <summary>
        /// Update an occurrence of a recurring meeting.
        /// </summary>
        /// <param name="occurrence">The occurrence to be updated.</param>
        /// <returns>If update operation succeeds, return true; otherwise, false.</returns>
        private bool UpdateOccurrenceItem(ItemType occurrence)
        {
            ItemIdType occurrenceId = occurrence.ItemId;

            if (occurrenceId != null)
            {
                CalendarItemType calendarUpdate = new CalendarItemType();
                calendarUpdate.Location = this.LocationUpdate;

                // Location change info
                AdapterHelper locationChangeInfo = new AdapterHelper();
                locationChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
                locationChangeInfo.Item = new CalendarItemType() { Location = this.LocationUpdate };
                locationChangeInfo.ItemId = occurrenceId;
                UpdateItemResponseMessageType itemOfLocationUpdate = this.UpdateSingleCalendarItem(Role.Organizer, locationChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
                Site.Assert.IsNotNull(itemOfLocationUpdate, "The location of the occurrence should be updated successfully.");
                CalendarItemType occurrenceOfLocationUpdate = itemOfLocationUpdate.Items.Items[0] as CalendarItemType;

                CalendarItemType calendar = occurrence as CalendarItemType;
                Site.Assert.IsNotNull(calendar, "The type conversion from ItemType to CalendarItemType should succeed.");

                // Start time change info
                DateTime start = calendar.Start;
                AdapterHelper startChangeInfo = new AdapterHelper();
                startChangeInfo.FieldURI = UnindexedFieldURIType.calendarStart;
                startChangeInfo.Item = new CalendarItemType() { Start = start.AddHours(-26.0), StartSpecified = true };
                startChangeInfo.ItemId = occurrenceOfLocationUpdate.ItemId;
                UpdateItemResponseMessageType itemOfStartUpdate = this.UpdateSingleCalendarItem(Role.Organizer, startChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
                Site.Assert.IsNotNull(itemOfStartUpdate, "The start time of the occurrence should be updated successfully.");
                CalendarItemType occurrenceOfStartUpdate = itemOfStartUpdate.Items.Items[0] as CalendarItemType;

                // End time change info
                DateTime end = calendar.End;
                AdapterHelper endChangeInfo = new AdapterHelper();
                endChangeInfo.FieldURI = UnindexedFieldURIType.calendarEnd;
                endChangeInfo.Item = new CalendarItemType() { End = end.AddHours(-26.0), EndSpecified = true };
                endChangeInfo.ItemId = occurrenceOfStartUpdate.ItemId;
                UpdateItemResponseMessageType itemOfEndUpdate = this.UpdateSingleCalendarItem(Role.Organizer, endChangeInfo, CalendarItemUpdateOperationType.SendOnlyToAll);
                Site.Assert.IsNotNull(itemOfEndUpdate, "The end time of the occurrence should be updated successfully.");

                return true;
            }

            return false;
        }

        /// <summary>
        /// Get the first occurrence of a recurring meeting.
        /// </summary>
        /// <param name="meetingItem">A recurring meeting.</param>
        /// <param name="role">The role to get the recurring meeting.</param>
        /// <returns>The first occurrence of the recurring meeting.</returns>
        private ItemType GetFirstOccurrenceItem(CalendarItemType meetingItem, Role role)
        {
            if (meetingItem != null)
            {
                #region Get the calendar the targeted occurrence belongs to.
                CalendarItemType calendar = this.SearchSingleItem(role, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
                Site.Assert.IsNotNull(calendar, "The calendar the targeted occurrence belongs to should exist.");

                OccurrenceItemIdType occurrenceId = new OccurrenceItemIdType();
                occurrenceId.RecurringMasterId = calendar.ItemId.Id;
                occurrenceId.InstanceIndex = 1;
                #endregion

                #region Get the occurrence item
                ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(role, occurrenceId);
                if (getItem != null)
                {
                    return getItem.Items.Items[0];
                }
                #endregion
            }

            return null;
        }

        /// <summary>
        /// Define the configuration of a recurring meeting.
        /// </summary>
        /// <param name="basePattern">A recurrence pattern.</param>
        /// <param name="range">A recurrence range.</param>
        /// <returns>A calendar item to be created.</returns>
        private CalendarItemType DefineRecurringMeeting(RecurrencePatternBaseType basePattern, RecurrenceRangeBaseType range)
        {
            CalendarItemType meetingItem = null;
            if (basePattern != null && range != null)
            {
                meetingItem = new CalendarItemType();

                // Define common property.
                meetingItem.UID = Guid.NewGuid().ToString();
                meetingItem.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
                meetingItem.Location = this.Location;
                meetingItem.IsResponseRequested = false;
                meetingItem.IsResponseRequestedSpecified = true;
                meetingItem.ConferenceType = 1;
                meetingItem.ConferenceTypeSpecified = true;
                meetingItem.AllowNewTimeProposal = true;
                meetingItem.AllowNewTimeProposalSpecified = true;
                meetingItem.LegacyFreeBusyStatus = LegacyFreeBusyType.OOF;
                meetingItem.LegacyFreeBusyStatusSpecified = true;

                DateTime startTime = DateTime.UtcNow.AddHours(3);
                meetingItem.Start = startTime;
                meetingItem.StartSpecified = true;

                DateTime endTime = startTime.AddHours(1);
                meetingItem.End = endTime;
                meetingItem.EndSpecified = true;

                // Set recurrence with specified pattern and range values.
                RecurrenceType recurrence = new RecurrenceType();
                AbsoluteYearlyRecurrencePatternType patternAbsoluteYearlyRecurrence = null;

                IntervalRecurrencePatternBaseType patternIntervalRecurrence = basePattern as IntervalRecurrencePatternBaseType;
                if (patternIntervalRecurrence != null)
                {
                    // Set the pattern's Interval.
                    patternIntervalRecurrence.Interval = this.PatternInterval;
                    recurrence.Item = patternIntervalRecurrence;
                }
                else
                {
                    patternAbsoluteYearlyRecurrence = basePattern as AbsoluteYearlyRecurrencePatternType;
                    if (patternAbsoluteYearlyRecurrence != null)
                    {
                        recurrence.Item = patternAbsoluteYearlyRecurrence;
                    }
                    else
                    {
                        RelativeYearlyRecurrencePatternType patternRelativeYearlyRecurrence = basePattern as RelativeYearlyRecurrencePatternType;
                        if (patternRelativeYearlyRecurrence != null)
                        {
                            recurrence.Item = patternRelativeYearlyRecurrence;
                        }
                    }
                }

                // Set the range's StartDate.
                DateTime startDate = startTime.AddMonths(1);
                range.StartDate = new DateTime(startDate.Year, startDate.Month, startDate.Day, 0, 0, 0, DateTimeKind.Utc);

                EndDateRecurrenceRangeType endDateRange = range as EndDateRecurrenceRangeType;
                if (endDateRange != null)
                {
                    if (patternAbsoluteYearlyRecurrence != null)
                    {
                        endDateRange.EndDate = range.StartDate.AddYears(8);
                    }
                    else
                    {
                        endDateRange.EndDate = range.StartDate.AddMonths(8);
                    }

                    recurrence.Item1 = endDateRange;
                }
                else
                {
                    recurrence.Item1 = range;
                }

                meetingItem.Recurrence = recurrence;

                meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
                meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            }

            return meetingItem;
        }

        /// <summary>
        /// Verify the value of CalendarItemCreateOrDeleteOperationType.
        /// </summary>
        /// <param name="calendarItemCreateOrDeleteOperationType">Specify a value of CalendarItemCreateOrDeleteOperationType;</param>
        private void VerifyCalendarItemCreateOrDeleteOperationType(CalendarItemCreateOrDeleteOperationType calendarItemCreateOrDeleteOperationType)
        {
            #region Step1: Organizer set the properties of the meeting to create
            CalendarItemType meeting = new CalendarItemType();
            meeting.UID = Guid.NewGuid().ToString();
            meeting.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));

            meeting.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meeting.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            #endregion

            #region Step2: Organizer creates the meeting
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meeting, calendarItemCreateOrDeleteOperationType);
            Site.Assert.IsNotNull(item, "Create a meeting item should be successful.");
            ItemIdType meetingId = item.Items.Items[0].ItemId;
            #endregion

            #region Step3: Verify CalendarItemCreateOrDeleteOperationType used in CreateItem operation
            #region find the message in Organizer Calendar
            bool createdIsFoundInOrgnizerCalendar = false;
            if (null != this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID))
            {
                createdIsFoundInOrgnizerCalendar = true;
            }
            #endregion

            #region find the message in Organizer SentItems
            bool createdIsFoundInOrgnizerSentItems = false;
            if (null != this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Request", meeting.UID))
            {
                createdIsFoundInOrgnizerSentItems = true;
            }
            #endregion

            #region find the message in Attendee Inbox
            bool createdIsFoundInAttendeeInbox = false;
            if (null != this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meeting.UID))
            {
                createdIsFoundInAttendeeInbox = true;
            }
            #endregion

            #region find the message in Attendee Calendar
            bool createdIsFoundInAttendeeCalendars = false;
            if (null != this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID))
            {
                createdIsFoundInAttendeeCalendars = true;
            }
            #endregion

            #region Verify relevant requirements for CalendarItemCreateOrDeleteOperationType used in CreateItem operation
            switch (calendarItemCreateOrDeleteOperationType)
            {
                case CalendarItemCreateOrDeleteOperationType.SendOnlyToAll:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R46");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInOrgnizerCalendar,
                        46,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: For the CreateItem operation, this value specifies that the meeting is created in the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4600");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInAttendeeInbox,
                        4600,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: [For the CreateItem operation, this value specifies that] a meeting request is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4601");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInAttendeeCalendars,
                        4601,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: [For the CreateItem operation, this value specifies that] the meeting is created in each attendee's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4602");

                    this.Site.CaptureRequirementIfIsFalse(
                        createdIsFoundInOrgnizerSentItems,
                        4602,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: [For the CreateItem operation, this value specifies that] No copy of the meeting request is saved in the organizer's Sent Items folder.");
                    break;

                case CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R48");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInOrgnizerCalendar,
                        48,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: For the CreateItem operation, this value specifies that the meeting is created in the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4800");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInAttendeeInbox,
                        4800,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: [For the CreateItem operation, this value specifies that] a meeting request is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4801");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInAttendeeCalendars,
                        4801,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: [For the CreateItem operation, this value specifies that] the meeting is created in each attendee's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4802");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInOrgnizerSentItems,
                        4802,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: [For the CreateItem operation, this value specifies that] A copy of the meeting request is saved in the organizer's Sent Items folder.");
                    break;

                case CalendarItemCreateOrDeleteOperationType.SendToNone:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R44");

                    this.Site.CaptureRequirementIfIsTrue(
                        createdIsFoundInOrgnizerCalendar,
                        44,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToNone: For the CreateItem operation ([MS-OXWSCORE] section 3.1.4.2), this value specifies that the meeting is created in the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4400");

                    this.Site.CaptureRequirementIfIsFalse(
                        createdIsFoundInAttendeeInbox,
                        4400,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToNone: [For the CreateItem operation ([MS-OXWSCORE] section 3.1.4.2), this value specifies that] no meeting request is sent to attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4401");

                    this.Site.CaptureRequirementIfIsFalse(
                        createdIsFoundInAttendeeCalendars,
                        4401,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToNone: [For the CreateItem operation ([MS-OXWSCORE] section 3.1.4.2), this value specifies that the meeting is created in the organizer's Calendar folder but no meeting request is sent to attendees.] Because no meeting request is generated, the meeting is not created in each attendee's Calendar folder.");
                    break;
            }
            #endregion
            #endregion

            #region Step4: Organizer delete the calendar item
            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, meetingId, calendarItemCreateOrDeleteOperationType);
            Site.Assert.IsNotNull(deletedItem, "Organizer should delete the calendar item successfully.");
            #endregion

            #region Step5: verify CalendarItemCreateOrDeleteOperationType used in DeleteItem operation
            #region find the message in Organizer Calendar
            bool deletedIsFoundInOrgnizerCalendar = false;

            if (null != this.SearchDeletedSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meeting.UID))
            {
                deletedIsFoundInOrgnizerCalendar = true;
            }
            #endregion

            #region find the message in Organizer SentItems
            bool deletedIsFoundInOrgnizerSentItems = false;
            if (null != this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.sentitems, "IPM.Schedule.Meeting.Canceled", meeting.UID))
            {
                deletedIsFoundInOrgnizerSentItems = true;
            }
            #endregion

            #region find the message in Attendee Inbox
            bool deletedIsFoundInAttendeeInbox = false;
            if (null != this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meeting.UID))
            {
                deletedIsFoundInAttendeeInbox = true;
            }
            #endregion

            #region Verify relevant requirements for CalendarItemCreateOrDeleteOperationType used in DeleteItem operation
            switch (calendarItemCreateOrDeleteOperationType)
            {
                case CalendarItemCreateOrDeleteOperationType.SendOnlyToAll:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R47");

                    this.Site.CaptureRequirementIfIsFalse(
                        deletedIsFoundInOrgnizerCalendar,
                        47,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: For the DeleteItem operation, this value specifies that the meeting is deleted from the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4700");

                    this.Site.CaptureRequirementIfIsTrue(
                        deletedIsFoundInAttendeeInbox,
                        4700,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: [For the DeleteItem operation, this value specifies that the meeting is deleted from the organizer's Calendar folder and] a meeting cancellation message is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4701");

                    this.Site.CaptureRequirementIfIsFalse(
                        deletedIsFoundInOrgnizerSentItems,
                        4701,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendOnlyToAll: [For the DeleteItem operation, this value specifies that] no copy of the meeting cancellation message is saved.");
                    break;

                case CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R49");

                    this.Site.CaptureRequirementIfIsFalse(
                        deletedIsFoundInOrgnizerCalendar,
                        49,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: For the DeleteItem operation, this value specifies that the meeting is deleted from the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4900");

                    this.Site.CaptureRequirementIfIsTrue(
                        deletedIsFoundInAttendeeInbox,
                        4900,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: [For the DeleteItem operation, this value specifies that] a meeting cancellation message is sent to all attendees.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4901");

                    this.Site.CaptureRequirementIfIsTrue(
                        deletedIsFoundInOrgnizerSentItems,
                        4901,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToAllAndSaveCopy: [For the DeleteItem operation, this value specifies that] a copy of the meeting cancellation message is saved in the organizer's Sent Items folder.");
                    break;

                case CalendarItemCreateOrDeleteOperationType.SendToNone:
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R45");

                    this.Site.CaptureRequirementIfIsFalse(
                        deletedIsFoundInOrgnizerCalendar,
                        45,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToNone: For the DeleteItem operation ([MS-OXWSCORE] section 3.1.4.3) this value specifies that the meeting is deleted from the organizer's Calendar folder.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R4500");

                    this.Site.CaptureRequirementIfIsFalse(
                        deletedIsFoundInAttendeeInbox,
                        4500,
                        @"[In t:CalendarItemCreateOrDeleteOperationType Simple Type] SendToNone: [For the DeleteItem operation ([MS-OXWSCORE] section 3.1.4.3) this value specifies that] no meeting cancellation message is sent to attendees.");
                    break;
            }
            #endregion
            #endregion
        }
        #endregion
    }
}