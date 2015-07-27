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
    /// This scenario is designed to test operations related to creation, retrieving, updating, copy, movement and deletion of multiple calendar related items on server.
    /// </summary>
    [TestClass]
    public class S05_OperateMultipleCalendarRelatedItems : TestSuiteBase
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
        /// This test case is designed to test getting multiple calendar items successfully. 
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S05_TC01_GetMultipleCalendarItems()
        {
            #region Define multiple calendar items
            int timeInterval = this.TimeInterval;
            CalendarItemType calendarItem1 = new CalendarItemType();
            calendarItem1.UID = Guid.NewGuid().ToString();
            calendarItem1.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            calendarItem1.Start = DateTime.Now.AddHours(timeInterval);

            // Indicate the Start property is serialized in the SOAP message.
            calendarItem1.StartSpecified = true;
            timeInterval++;
            calendarItem1.End = DateTime.Now.AddHours(timeInterval);
            calendarItem1.EndSpecified = true;
            calendarItem1.LegacyFreeBusyStatus = this.LegacyFreeBusy;
            calendarItem1.LegacyFreeBusyStatusSpecified = true;
            calendarItem1.Location = this.Location;
            calendarItem1.When = string.Format("{0} to {1}", calendarItem1.Start.ToString(), calendarItem1.End.ToString());

            CalendarItemType calendarItem2 = new CalendarItemType();
            calendarItem2.UID = Guid.NewGuid().ToString();
            calendarItem2.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            timeInterval = this.TimeInterval;
            calendarItem2.Start = calendarItem1.End.AddHours(timeInterval);
            timeInterval++;
            calendarItem2.StartSpecified = true;
            calendarItem2.End = calendarItem1.End.AddHours(timeInterval);
            calendarItem2.EndSpecified = true;
            calendarItem2.LegacyFreeBusyStatus = this.LegacyFreeBusy;
            calendarItem2.LegacyFreeBusyStatusSpecified = true;
            calendarItem2.Location = this.Location;
            calendarItem2.When = string.Format("{0} to {1}", calendarItem1.Start.ToString(), calendarItem1.End.ToString());
            #endregion

            #region Create multiple calendar items
            ItemInfoResponseMessageType[] calendars = this.CreateMultipleCalendarItems(Role.Organizer, new ItemType[] { calendarItem1, calendarItem2 }, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(calendars, "The calendars should be created successfully.");
            Site.Assert.IsTrue(calendars.Length == 2, "There should be only two calendars created.");

            ItemIdType[] calendarIds = new ItemIdType[] { calendars[0].Items.Items[0].ItemId, calendars[1].Items.Items[0].ItemId };
            #endregion

            #region Get multiple calendar items
            ItemInfoResponseMessageType[] getItems = this.GetMultipleCalendarItems(Role.Organizer, calendarIds);
            Site.Assert.IsNotNull(getItems, "The calendars should be gotten successfully.");
            Site.Assert.IsTrue(getItems.Length == 2, "There should be only two calendars returned by GetItem.");
            #endregion

            #region Delete multiple calendar items
            Site.Assert.IsNotNull(
                this.DeleteMultipleCalendarItems(Role.Organizer, calendarIds, CalendarItemCreateOrDeleteOperationType.SendToNone),
                "Organizer should delete multiple calendar items successfully.");
            #endregion
        }

        /// <summary>GenerateResourceName(this.Site
        /// This test case is designed to test updating multiple calendar items successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S05_TC02_UpdateMultipleCalendarItems()
        {
            #region Define two calendar items
            CalendarItemType calendarItem1 = new CalendarItemType();
            calendarItem1.UID = Guid.NewGuid().ToString();
            calendarItem1.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            calendarItem1.Location = this.Location;

            CalendarItemType calendarItem2 = new CalendarItemType();
            calendarItem2.UID = Guid.NewGuid().ToString();
            calendarItem2.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            calendarItem2.Location = this.Location;
            #endregion

            #region Create the two calendar items
            ItemInfoResponseMessageType[] calendars = this.CreateMultipleCalendarItems(Role.Organizer, new ItemType[] { calendarItem1, calendarItem2 }, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(calendars, "The calendars should be created successfully.");
            Site.Assert.IsTrue(calendars.Length == 2, "There should be only two calendars created.");

            ItemIdType[] calendarIds = new ItemIdType[] { calendars[0].Items.Items[0].ItemId, calendars[1].Items.Items[0].ItemId };
            #endregion

            #region Update the Location element of the two created calendar items
            List<AdapterHelper> itemsChangeInfo = new List<AdapterHelper>();
            foreach (ItemIdType calendarId in calendarIds)
            {
                CalendarItemType calendarUpdate = new CalendarItemType();
                calendarUpdate.Location = this.LocationUpdate;

                AdapterHelper itemChangeInfo = new AdapterHelper();
                itemChangeInfo.FieldURI = UnindexedFieldURIType.calendarLocation;
                itemChangeInfo.Item = calendarUpdate;
                itemChangeInfo.ItemId = calendarId;
                itemsChangeInfo.Add(itemChangeInfo);
            }

            Site.Assert.IsNotNull(
                this.UpdateMultipleCalendarItems(Role.Organizer, itemsChangeInfo.ToArray(), CalendarItemUpdateOperationType.SendToNone),
                "Server should return success for updating multiple calendar items.");
            #endregion

            #region Verify the Location elements of the two calendar items are updated
            ItemInfoResponseMessageType getItem1 = this.GetSingleCalendarItem(Role.Organizer, calendarIds[0]);
            Site.Assert.IsNotNull(getItem1, "The first updated item should exist.");

            CalendarItemType calendar1 = getItem1.Items.Items[0] as CalendarItemType;
            Site.Assert.AreEqual<string>(
                this.LocationUpdate,
                calendar1.Location,
                string.Format("The Location of the first updated calendar should be {0}. The actual value is {1}.", this.LocationUpdate, calendar1.Location));

            ItemInfoResponseMessageType getItem2 = this.GetSingleCalendarItem(Role.Organizer, calendarIds[1]);
            Site.Assert.IsNotNull(getItem2, "The second updated item should exist.");

            CalendarItemType calendar2 = getItem2.Items.Items[0] as CalendarItemType;
            Site.Assert.AreEqual<string>(
                this.LocationUpdate,
                calendar2.Location,
                string.Format("The Location of the second updated calendar should be {0}. The actual value is {1}.", this.LocationUpdate, calendar2.Location));
            #endregion

            #region Clean up organizer's calendar folder.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test copying multiple calendar items successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S05_TC03_CopyMultipleCalendarItems()
        {
            #region Define two calendar items to copy
            CalendarItemType calendarItem1 = new CalendarItemType();
            calendarItem1.UID = Guid.NewGuid().ToString();
            calendarItem1.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            CalendarItemType calendarItem2 = new CalendarItemType();
            calendarItem2.UID = Guid.NewGuid().ToString();
            calendarItem2.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            #endregion

            #region Organizer creates the two calendar items
            ItemInfoResponseMessageType[] calendars = this.CreateMultipleCalendarItems(Role.Organizer, new ItemType[] { calendarItem1, calendarItem2 }, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(calendars, "The calendars should be created successfully.");
            Site.Assert.IsTrue(calendars.Length == 2, "There should be only two calendars created.");

            ItemIdType[] calendarIds = new ItemIdType[] { calendars[0].Items.Items[0].ItemId, calendars[1].Items.Items[0].ItemId };
            #endregion

            #region Organizer copies the two calendar items to Drafts folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.drafts;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            Site.Assert.IsNotNull(
                this.CopyMultipleCalendarItems(Role.Organizer, calendarIds, targetFolderId),
                "The items should be copied successfully.");
            #endregion

            #region Organizer calls GetItem operation to verify whether the calendar items are really copied
            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem1.UID),
                "The original calendar should be in organizer's calendar folder after CopyItem operation.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.drafts, "IPM.Appointment", calendarItem1.UID),
                "The original calendar should also be in organizer's drafts folder after CopyItem operation.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem2.UID),
                "The original calendar should be in organizer's calendar folder after CopyItem operation.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.drafts, "IPM.Appointment", calendarItem2.UID),
                "The original calendar should also be in organizer's drafts folder after CopyItem operation.");
            #endregion

            #region Clean up organizer's drafts and calendar folders.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.calendar, DistinguishedFolderIdNameType.drafts });
            #endregion
        }

        /// <summary>
        /// This test case is designed to test moving multiple calendar items successfully.
        /// </summary>
        [TestCategory("MSOXWSMTGS"), TestMethod()]
        public void MSOXWSMTGS_S05_TC04_MoveMultipleCalendarItems()
        {
            #region Define two calendar items to move
            CalendarItemType calendarItem1 = new CalendarItemType();
            calendarItem1.UID = Guid.NewGuid().ToString();
            calendarItem1.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            CalendarItemType calendarItem2 = new CalendarItemType();
            calendarItem2.UID = Guid.NewGuid().ToString();
            calendarItem2.Subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("MeetingSubject", this.Site));
            #endregion

            #region Create the two calendar items
            ItemInfoResponseMessageType[] calendars = this.CreateMultipleCalendarItems(Role.Organizer, new ItemType[] { calendarItem1, calendarItem2 }, CalendarItemCreateOrDeleteOperationType.SendToNone);
            Site.Assert.IsNotNull(calendars, "The calendars should be created successfully.");
            Site.Assert.IsTrue(calendars.Length == 2, "There should be only two calendars created.");

            ItemIdType[] calendarIds = new ItemIdType[] { calendars[0].Items.Items[0].ItemId, calendars[1].Items.Items[0].ItemId };
            #endregion

            #region Move the two calendar items to Inbox folder
            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            folderId.Id = DistinguishedFolderIdNameType.inbox;
            TargetFolderIdType targetFolderId = new TargetFolderIdType();
            targetFolderId.Item = folderId;

            Site.Assert.IsNotNull(
                this.MoveMultipleCalendarItems(Role.Organizer, calendarIds, targetFolderId),
                "The calendars should be moved into the inbox folder successfully.");

            #endregion

            #region Call FindItem to verify the two calendar items are moved to Inbox folder
            Site.Assert.IsNull(
                this.SearchDeletedSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem1.UID),
                "The original calendar should not be in organizer's calendar folder after MoveItem operation.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Appointment", calendarItem1.UID),
                "The original calendar should be in organizer's inbox folder after MoveItem operation.");

            Site.Assert.IsNull(
                this.SearchDeletedSingleItem(Role.Organizer, DistinguishedFolderIdNameType.calendar, "IPM.Appointment", calendarItem2.UID),
                "The original calendar should not be in organizer's calendar folder after MoveItem operation.");

            Site.Assert.IsNotNull(
                this.SearchSingleItem(Role.Organizer, DistinguishedFolderIdNameType.inbox, "IPM.Appointment", calendarItem2.UID),
                "The original calendar should be in organizer's inbox folder after MoveItem operation.");
            #endregion

            #region Clean up organizer's inbox folder.
            this.CleanupFoldersByRole(Role.Organizer, new List<DistinguishedFolderIdNameType>() { DistinguishedFolderIdNameType.inbox });
            #endregion
        }
        #endregion
    }
}