namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to copy of calendar related items on server.
    /// </summary>
    [TestClass]
    public class S03_CopyCalendarRelatedItem : TestSuiteBase
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
        /// This test case is designed to test copying a single appointment item(calendar item without attendeeType) successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S03_TC01_CopySingleCalendar()
        {
            #region Define a calendar item to copy
            CalendarItemType calendarItem = new CalendarItemType();
            calendarItem.UID = Guid.NewGuid().ToString();
            calendarItem.Subject = this.Subject;
            #endregion

            #region Create the calendar item with CalendarItemCreateOrDeleteOperationType set to SendToNone
            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, calendarItem, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(item, "Create a calendar item should be successful.");
            ItemIdType calendarId = item.Items.Items[0].ItemId;
            #endregion

            #region Copy the calendar item to Drafts folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.drafts;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            ItemInfoResponseMessageType copiedItem = this.CopySingleCalendarItem(Role.Organizer, calendarId, targetFolderId);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R602");

            // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R602
            this.Site.CaptureRequirementIfIsNotNull(
                copiedItem,
                602,
                @"[In Messages] CopyItemSoapIn: For each item being copied that is not a recurring calendar item, the ItemIds element MUST contain an ItemId child element ([MS-OXWSCORE] section 2.2.4.11).");
            #endregion

            #region Call GetItem operation to verify whether the calendar item is really copied
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.drafts, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The calendar item should be in organizer's drafts folder.");

            CalendarItemType calendarInCalendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendarInCalendar, "The calendar item should also be in organizer's calendar folder.");
            #endregion

            #region Clean up organizer's drafts and calendar folders
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.drafts });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test copying meeting item, meeting request message, meeting response message and meeting cancellation message successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S03_TC02_CopySingleMeetingItem()
        {
            #region Organizer creates a single meeting getItem with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            CalendarItemType meetingItem = new CalendarItemType();
            meetingItem.Subject = this.Subject;
            meetingItem.UID = Guid.NewGuid().ToString();
            meetingItem.RequiredAttendees = new AttendeeType[] { GetAttendeeOrResource(this.AttendeeEmailAddress) };
            meetingItem.OptionalAttendees = new AttendeeType[] { GetAttendeeOrResource(this.OrganizerEmailAddress) };
            meetingItem.Resources = new AttendeeType[] { GetAttendeeOrResource(this.RoomEmailAddress) };

            ItemInfoResponseMessageType item = this.CreateSingleCalendarItem(Role.Organizer, meetingItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(item, "Create single meeting item should be successful.");
            #endregion

            #region Organizer copies the created single meeting item to Drafts folder
            CalendarItemType calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The created calendar should exist in organizer's calendar folder.");
            ItemIdType itemId = calendar.ItemId;

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.drafts;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            ItemInfoResponseMessageType copiedItem = this.CopySingleCalendarItem(Role.Organizer, itemId, targetFolderId);
            Site.Assert.IsNotNull(copiedItem, @"Copy the single meeting item should be successful.");
            #endregion

            #region Organizer calls GetItem operation to verify whether the meeting item is really copied
            calendar = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.drafts, "IPM.Appointment", meetingItem.UID) as CalendarItemType;
            Site.Assert.IsNotNull(calendar, "The copied calendar should exist in organizer's Drafts folder.");
            #endregion

            #region Attendee gets the meeting request message in the Inbox folder
            MeetingRequestMessageType request = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, meetingItem.Subject, meetingItem.UID, UnindexedFieldURIType.itemSubject) as MeetingRequestMessageType;
            Site.Assert.IsNotNull(request, "The meeting request message should exist in attendee's inbox folder.");
            #endregion

            #region Attendee copies the meeting request message to the Drafts folder
            copiedItem = this.CopySingleCalendarItem(Role.Attendee, request.ItemId, targetFolderId);
            Site.Assert.IsNotNull(copiedItem, @"Copy the single meeting request message should be successful.");
            #endregion

            #region Attendee calls CreateItem to accept the meeting request with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            AcceptItemType acceptItem = new AcceptItemType();
            acceptItem.ReferenceItemId = request.ItemId;

            Site.Assert.IsNotNull(
                this.CreateSingleCalendarItem(Role.Attendee, acceptItem, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll),
                "Attendee creates items for meeting request should succeed.");
            #endregion

            #region Organizer finds the meeting response message in his Inbox folder and copies it to the Drafts folder
            MeetingResponseMessageType response = this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Resp", meetingItem.UID) as MeetingResponseMessageType;
            Site.Assert.IsNotNull(response, "The response message from Attendee should be in organizer's Inbox folder.");

            copiedItem = this.CopySingleCalendarItem(Role.Organizer, response.ItemId, targetFolderId);
            Site.Assert.IsNotNull(copiedItem, @"Copy the single meeting response message should be successful.");
            #endregion

            #region Organizer deletes the meeting item with CalendarItemCreateOrDeleteOperationType value set to SendOnlyToAll
            ResponseMessageType deletedItem = this.DeleteSingleCalendarItem(Role.Organizer, itemId, CalendarItemCreateOrDeleteOperationType.SendOnlyToAll);
            Site.Assert.IsNotNull(deletedItem, @"Delete the single meeting item should be successful.");
            #endregion

            #region Attendee finds the meeting cancellation message in the Inbox folder and copies it to the Drafts folder
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int count = 1;

            MeetingCancellationMessageType canceledItem = null;

            while (canceledItem == null && count++ <= upperBound)
            {
                canceledItem = this.SearchSingleItem(Role.Attendee, DistinguishedFolderIdNameType.inbox, "IPM.Schedule.Meeting.Canceled", meetingItem.UID) as MeetingCancellationMessageType;
                System.Threading.Thread.Sleep(waitTime);
            }

            Site.Assert.IsNotNull(canceledItem, "The cancellation meeting message should be in attendee's Inbox folder.");

            ItemIdType canceledItemId = canceledItem.ItemId;
            copiedItem = this.CopySingleCalendarItem(Role.Attendee, canceledItemId, targetFolderId);
            Site.Assert.IsNotNull(copiedItem, "Attendee should copy the meeting cancellation message to the Drafts folder.");
            #endregion

            #region Clean up inbox, drafts and deleteditems folders of both organizer and attendee. Attendee's sentitems and calendar should also be cleaned up.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.drafts, DistinguishedFolderIdNameType.deleteditems });
            this.CleanupFoldersByRole(Role.Attendee, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox, DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.sentitems, DistinguishedFolderIdNameType.drafts, DistinguishedFolderIdNameType.deleteditems });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test copying a recurring calendar item successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S03_TC03_CopyRecurringCalendar()
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

            #region Get the targeted occurrence item
            ItemInfoResponseMessageType getItem = this.GetSingleCalendarItem(Role.Organizer, occurrenceItemId);
            Site.Assert.IsNotNull(getItem, @"Get the occurrence should be successful.");

            RecurringMasterItemIdType recurringMasterItemId = new RecurringMasterItemIdType();
            recurringMasterItemId.ChangeKey = getItem.Items.Items[0].ItemId.ChangeKey;
            recurringMasterItemId.OccurrenceId = getItem.Items.Items[0].ItemId.Id;
            #endregion

            #region Copy the recurring calendar item to Drafts folder through Id of recurring master getItem
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.drafts;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            ItemInfoResponseMessageType copiedItem = this.CopySingleCalendarItem(Role.Organizer, recurringMasterItemId, targetFolderId);
            Site.Assert.IsNotNull(copiedItem, @"Copy recurring calendar item through RecurringMasterItemId should be successful.");
            ItemIdType calendarIdByCopied = copiedItem.Items.Items[0].ItemId;
            #endregion

            #region Call GetItem operation to verify whether the recurring calendar item is really copied
            getItem = this.GetSingleCalendarItem(Role.Organizer, calendarIdByCopied);

            if (Common.IsRequirementEnabled(806, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R806");

                // Verify MS-OXWSMTGS requirement: MS-OXWSMTGS_R806
                this.Site.CaptureRequirementIfIsNotNull(
                    getItem,
                    806,
                    @"[In Appendix C: Product Behavior] CopyItemSoapIn: For each item being copied that is a recurring calendar item, implementation does contain a RecurringMasterItemId child element ([MS-OXWSCORE] section 2.2.4.11). (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Clean up organizer's drafts and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.drafts });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test ErrorCalendarCannotMoveOrCopyOccurrence will be returned if copy an occurrence of a recurring calendar item.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S03_TC04_CopyItemErrorCalendarCannotMoveOrCopyOccurrence()
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

            CopyItemType copyItemRequest = new CopyItemType();
            copyItemRequest.ItemIds = new BaseItemIdType[] { occurrenceItemId };
            copyItemRequest.ToFolderId = targetFolderId;
            CopyItemResponseType response = this.MTGSAdapter.CopyItem(copyItemRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1190");

            // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1190
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                response.ResponseMessages.Items[0].ResponseClass,
                1190,
                @"[In Messages] If the request is unsuccessful, the CopyItem operation returns an CopyItemResponse element with the ResponseClass attribute of the CopyItemResponseMessage element set to ""Error"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMTGS_R1193");

            // Verify MS-OXWSMSG requirement: MS-OXWSMTGS_R1193
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCalendarCannotMoveOrCopyOccurrence,
                response.ResponseMessages.Items[0].ResponseCode,
                1193,
                @"[In Messages] ErrorCalendarCannotMoveOrCopyOccurrence: Specifies that an attempt was made to move or copy an occurrence of a recurring calendar item.");
            #endregion

            #region Clean up organizer's calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            #endregion
        }
        #endregion
    }
}