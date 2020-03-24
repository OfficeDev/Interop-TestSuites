namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to movement of calendar related items on server.
    /// </summary>
    [TestClass]
    public class S04_MoveCalendarRelatedItem : TestSuiteBase
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
        /// This test case is designed to test moving a single appointment item(calendar item without attendeeType) successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S04_TC01_MoveSingleCalendar()
        {
            #region Define a calendar item to move
            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            #endregion

            #region Create the calendar item
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            ItemIdType calendarId = item.Items.Items[0].ItemId;
            #endregion

            #region Move the created calendar item to Inbox folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            ItemInfoResponseMessageType movedItem = this.MoveSingleCalendarItem(Role.Organizer, calendarId, targetFolderId);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R640");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R640
            this.Site.CaptureRequirementIfIsNotNull(
                movedItem,
                640,
                @"[In Messages] MoveItemSoapIn: For each item being moved that is not a recurring calendar item, the ItemIds element MUST contain an ItemId child element ([MS-OXWSCORE] section 2.2.4.11).");
            #endregion

            #region Verify the calendar item is moved to Inbox folder
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar item should be moved to Inbox folder.");

            CalendarItemType calendarInCalendar = this.SearchDeletedSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNull(calendarInCalendar, "The calendar item should not be in the Calendar folder.");
            #endregion

            #region Clean up organizer's inbox folder.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test moving meeting item, meeting request message, meeting response message and meeting cancellation message successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S04_TC02_MoveMeeting()
        {
            #region Define the target folder to move
            // Define the Inbox folder as the target folder for moving meeting item.
            DistinguishedFolderIdType inboxFolder = new DistinguishedFolderIdType();
            inboxFolder.Id = DistinguishedFolderIdNameType.inbox;
            TargetFolderIdType meetingTargetFolder = new TargetFolderIdType();
            meetingTargetFolder.Item = inboxFolder;

            // Define the Calendar folder as the target folder for moving meeting request message, meeting response message and meeting cancellation message.
            DistinguishedFolderIdType calendarFolder = new DistinguishedFolderIdType();
            calendarFolder.Id = DistinguishedFolderIdNameType.calendar;
            TargetFolderIdType msgTargetFolder = new TargetFolderIdType();
            msgTargetFolder.Item = calendarFolder;
            #endregion

            #region Organizer creates a meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            #region Define a meeting to be moved
            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.Subject = this.Subject;
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };
            #endregion

            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            ItemIdType meetingId = item.Items.Items[0].ItemId;
            #endregion

            #region Organizer moves the meeting item to Inbox folder
            ItemInfoResponseMessageType movedItem = this.MoveSingleCalendarItem(Role.Organizer, meetingId, meetingTargetFolder);

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                movedItem.ResponseClass,
                "Server should return success for moving the meeting item.");

            ItemIdType movedMeetingItemId = movedItem.Items.Items[0].ItemId;
            #endregion

            #region Organizer calls FindItem to verify the meeting item is moved to Inbox folder
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The meeting item should be moved into organizer's Inbox folder.");
            #endregion

            #region Attendee finds the meeting request message in his Inbox folder
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "The meeting request message should exist in attendee's inbox folder.");
            #endregion

            #region Attendee moves the meeting request message to the Calendar folder
            movedItem = this.MoveSingleCalendarItem(Role.Attendee, request.ItemId, msgTargetFolder);

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                movedItem.ResponseClass,
                @"Server should return success for moving meeting request message.");
            #endregion

            #region Attendee calls FindItem to verify the meeting request message is moved to Calendar folder
            request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Schedule.Meeting.Request", meetingItem.UID) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "The meeting request message should exist in attendee's calendar folder.");
            #endregion

            #region Attendee calls CreateItem to accept the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            AcceptItemType acceptItem = new AcceptItemType();
            acceptItem.ReferenceItemId = request.ItemId;

            item = this.CreateSingleCalendarItem(Role.Attendee, acceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            #endregion

            #region Organizer finds the meeting response message in its Inbox folder
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The meeting response message should exist in organizer's inbox folder.");
            #endregion

            #region Organizer moves the meeting response message to his Calendar folder
            movedItem = this.MoveSingleCalendarItem(Role.Organizer, response.ItemId, msgTargetFolder);

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                movedItem.ResponseClass,
                "EWS should return success for moving the meeting response message.");
            #endregion

            #region Organizer calls FindItem to verify the meeting response message is moved to Calendar folder
            response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The meeting response message should be in organizer's calendar folder.");
            #endregion

            #region Organizer calls DeleteItem to delete the meeting with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, movedMeetingItemId, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(deletedItem, "Delete the meeting should be successful.");
            #endregion

            #region Attendee finds the meeting cancellation message in his Inbox folder
            MeetingCancellationMessageType canceledMeeting = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNotNull(canceledMeeting, "The meeting cancellation message should be in attendee's inbox folder.");
            #endregion

            #region Attendee moves the meeting cancellation message to his Calendar folder
            movedItem = this.MoveSingleCalendarItem(Role.Attendee, canceledMeeting.ItemId, msgTargetFolder);

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Success,
                movedItem.ResponseClass,
                "EWS should return success for moving the meeting cancellation message.");
            #endregion

            #region Attendee calls FindItem to verify the meeting cancellation message is moved to Calendar folder
            canceledMeeting = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.calendar, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
            Site.Assert.IsNotNull(canceledMeeting, "The meeting cancellation message should be in attendee's calendar folder.");
            #endregion

            #region Attendee removes the meeting item
            RemoveItemType removeItem = new RemoveItemType();
            removeItem.ReferenceItemId = canceledMeeting.ItemId;
            item = this.CreateSingleCalendarItem(Role.Attendee, removeItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "The meeting item should be removed.");
            #endregion

            #region Clean up organizer's calendar and deleteditems folders, and attendee's sentitems and deleteditems folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test moving a recurring calendar item successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S04_TC03_MoveRecurringCalendar()
        {
            #region Define a recurring calendar item to move
            DateTime startTime = DateTime.Now;

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = this.PatternInterval;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = this.NumberOfOccurrences;
            range.StartDate = startTime;

            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.Start = startTime;
            calendarItem.StartSpecified = true;
            calendarItem.End = startTime.AddHours(this.TimeInterval);
            calendarItem.EndSpecified = true;
            calendarItem.Recurrence = new RecurrenceType();
            calendarItem.Recurrence.Item = pattern;
            calendarItem.Recurrence.Item1 = range;
            #endregion

            #region Create the recurring calendar item and extract the Id of an occurrence item
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            OccurrenceItemIdType occurrenceItemId = new OccurrenceItemIdType();
            occurrenceItemId.ChangeKey = item.Items.Items[0].ItemId.ChangeKey;
            occurrenceItemId.RecurringMasterId = item.Items.Items[0].ItemId.Id;
            occurrenceItemId.InstanceIndex = this.InstanceIndex;
            #endregion

            #region Get the Id of the occurrence item
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, occurrenceItemId);
            Site.Assert.IsNotNull(getItem, "Organizer should get the occurrence item successfully.");

            RecurringMasterItemIdType recurringMasterItemId = new RecurringMasterItemIdType();
            recurringMasterItemId.ChangeKey = getItem.Items.Items[0].ItemId.ChangeKey;
            recurringMasterItemId.OccurrenceId = getItem.Items.Items[0].ItemId.Id;
            #endregion

            #region Move the recurring calendar item to Inbox folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            ItemInfoResponseMessageType movedItem = this.MoveSingleCalendarItem(Role.Organizer, recurringMasterItemId, targetFolderId);
            Site.Assert.IsNotNull(movedItem, @"Server should return success for moving the recurring calendar item.");
            #endregion

            #region Call FindItem to verify the recurring calendar item is moved to Inbox folder
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The recurring calendar should be in organizer's inbox folder.");

            if (Common.IsRequirementEnabled(808, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R808");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R808
                this.Site.CaptureRequirementIfAreEqual<CalendarItemTypeType>(
                    CalendarItemTypeType.RecurringMaster,
                    calendar.CalendarItemType1,
                    808,
                    @"[In Appendix C: Product Behavior] MoveItemSoapIn: For each item being moved that is a recurring calendar item, implementation does contain a RecurringMasterItemId child element ([MS-OXWSCORE] section 2.2.4.11). (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Clean up organizer's inbox folder.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test ErrorCalendarCannotMoveOrCopyOccurrence will be returned if move an occurrence of a recurring calendar item.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S04_TC04_MoveItemErrorCalendarCannotMoveOrCopyOccurrence()
        {
            #region Define a recurring calendar item
            DateTime startTime = DateTime.Now;

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = this.PatternInterval;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = this.NumberOfOccurrences;
            range.StartDate = startTime;

            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.Start = startTime;
            calendarItem.StartSpecified = true;
            calendarItem.End = startTime.AddHours(this.TimeInterval);
            calendarItem.EndSpecified = true;
            calendarItem.Recurrence = new RecurrenceType();
            calendarItem.Recurrence.Item = pattern;
            calendarItem.Recurrence.Item1 = range;
            #endregion

            #region Create the recurring calendar item and extract the Id of an occurrence item
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);

            OccurrenceItemIdType occurrenceItemId = new OccurrenceItemIdType();
            occurrenceItemId.ChangeKey = item.Items.Items[0].ItemId.ChangeKey;
            occurrenceItemId.RecurringMasterId = item.Items.Items[0].ItemId.Id;
            occurrenceItemId.InstanceIndex = this.InstanceIndex;
            #endregion

            #region Copy one occurrence of the recurring calendar item
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.drafts;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            MoveItemType moveItemRequest = new MoveItemType();
            moveItemRequest.ItemIds = new BaseItemIdType[] { occurrenceItemId };
            moveItemRequest.ToFolderId = targetFolderId;
            MoveItemResponseType response = this.MTGSAdapter.MoveItem(moveItemRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1228");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1228
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                response.ResponseMessages.Items[0].ResponseClass,
                1228,
                @"[In Messages] If the request is unsuccessful, the MoveItem operation returns a MoveItemResponse element with the ResponseClass attribute of the MoveItemResponseMessage element set to ""Error"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1231");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1231
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCalendarCannotMoveOrCopyOccurrence,
                response.ResponseMessages.Items[0].ResponseCode,
                1231,
                @"[In Messages] ErrorCalendarCannotMoveOrCopyOccurrence: Specifies that an attempt was made to move or copy an occurrence of a recurring calendar item.");
            #endregion

            #region Clean up organizer's calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test ErrorCalendarCannotUseIdForOccurrenceId will be returned if the OccurrenceId does not correspond to a valid occurrence of the recurringMasterId.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S04_TC05_MoveItemErrorCalendarCannotUseIdForOccurrenceId()
        {
            #region Define a meeting
            int timeInterval = this.TimeInterval;
            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.Subject = this.Subject;
            meetingItem.Start = DateTime.UtcNow;

            meetingItem.StartSpecified = true;
            timeInterval++;
            meetingItem.End = DateTime.Now.AddHours(timeInterval);
            meetingItem.EndSpecified = true;
            meetingItem.Location = this.Location;

            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };

            // Create the  meeting
            ItemInfoResponseMessageType itemInfo = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy);
            Site.Assert.IsNotNull(itemInfo, "Server should return success for creating a recurring meeting.");

            CalendarItemType calendarInOrganizer = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendarInOrganizer, "The meeting should be found in organizer's Calendar folder after organizer calls CreateItem with CalendarItemCreateOrDeleteOperationType set to SendToAllAndSaveCopy.");
            #endregion

            #region Define a recurring calendar item to move
            DateTime startTime = DateTime.Now;

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = this.PatternInterval;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = this.NumberOfOccurrences;
            range.StartDate = startTime;

            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            calendarItem.Start = startTime;
            calendarItem.StartSpecified = true;
            calendarItem.End = startTime.AddHours(this.TimeInterval);
            calendarItem.EndSpecified = true;
            calendarItem.Recurrence = new RecurrenceType();
            calendarItem.Recurrence.Item = pattern;
            calendarItem.Recurrence.Item1 = range;
            #endregion

            #region Create the recurring calendar item and extract the Id of an occurrence item
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            OccurrenceItemIdType occurrenceItemId = new OccurrenceItemIdType();
            occurrenceItemId.ChangeKey = item.Items.Items[0].ItemId.ChangeKey;
            occurrenceItemId.RecurringMasterId = item.Items.Items[0].ItemId.Id;
            occurrenceItemId.InstanceIndex = this.InstanceIndex;
            #endregion

            #region Get the Id of the occurrence item
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, occurrenceItemId);
            Site.Assert.IsNotNull(getItem, "Organizer should get the occurrence item successfully.");

            RecurringMasterItemIdType recurringMasterItemId = new RecurringMasterItemIdType();
            recurringMasterItemId.ChangeKey = getItem.Items.Items[0].ItemId.ChangeKey;
            recurringMasterItemId.OccurrenceId = calendarInOrganizer.ItemId.Id;
            #endregion

            #region Move the recurring calendar item to Inbox folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            MoveItemType moveItemRequest = new MoveItemType();
            moveItemRequest.ItemIds = new BaseItemIdType[] { recurringMasterItemId };
            moveItemRequest.ToFolderId = targetFolderId;
            MoveItemResponseType response = this.MTGSAdapter.MoveItem(moveItemRequest);
            #endregion

            #region Capture code
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1232");

            //Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R1232
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCalendarCannotUseIdForOccurrenceId,
                response.ResponseMessages.Items[0].ResponseCode,
                1232,
                @"[In Messages] ErrorCalendarCannotUseIdForOccurrenceId: Specifies that the OccurrenceId ([MS-OXWSCORE] section 2.2.4.39) does not correspond to a valid occurrence of a recurring master item.");
            #endregion

            #region Clean up organizer's calendar, sentitems and deleteditems folders and attendee's calendar, inbox and deleteditems folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType> { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }
        #endregion
    }
}