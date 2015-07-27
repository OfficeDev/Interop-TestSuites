//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A test class contains test cases of S02 scenario.
    /// </summary>
    [TestClass]
    public class S02_OperateListItems : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test cases

        #region MSOUTSPS_S02_TC01_OperationListItemsForAppointment
        /// <summary>
        /// This test case is used to verify Appointment template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC01_OperationListItemsForAppointment()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;
            recurrenceXMLData.recurrence.rule.Item = "7";
            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            // If EventType is "1", this property MUST contain a valid TimeZoneXML.
            TimeZoneXML customPacificTimeZone = this.GetCustomPacificTimeZoneXmlSetting();
            recurEventFieldsSetting.Add("XMLTZone", this.GetTimeZoneXMLString(customPacificTimeZone));

            // Setting Duration field's value
            recurEventFieldsSetting.Add("Duration", "0");
            recurEventFieldsSetting.Add("Title", eventTitle);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");
            string actualEventDateValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventDate");
            string actualEndDateValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EndDate");

            #region Capture code

            // If the EndDate and EventDate fields' values equal to the values the client set in upon steps, then capture R1264
            DateTime actualEndDate;
            if (!DateTime.TryParse(actualEndDateValue, out actualEndDate))
            {
                this.Site.Assert.Fail("The EndDate field value should be a valid DateTime format.");
            }

            DateTime actualEventDate;
            if (!DateTime.TryParse(actualEventDateValue, out actualEventDate))
            {
                this.Site.Assert.Fail("The EventDate field value should be a valid DateTime format.");
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R267
            this.Site.CaptureRequirementIfIsTrue(
                DateTime.Parse(actualEndDateValue) >= DateTime.Parse(actualEventDateValue),
                267,
                "[In Appointment-Specific Schema]EndDate MUST be equal to or later than EventDate.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R271
            this.Site.CaptureRequirementIfIsTrue(
                DateTime.Parse(actualEndDateValue) >= DateTime.Parse(actualEventDateValue),
                271,
                "[In Appointment-Specific Schema]EventDate MUST be equal to or earlier than EndDate.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R272
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(actualEventDateValue),
                272,
                "[In Appointment-Specific Schema]EventDate MUST NOT be empty or missing.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R275
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType")),
                275,
                "[In Appointment-Specific Schema]EventType MUST NOT be empty or missing.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R288
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence")),
                288,
                "[In Appointment-Specific Schema]fRecurrence MUST NOT be empty or missing.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC02_OperationListItems_fAllDayEvent
        /// <summary>
        /// This test case is used to verify if the fAllDayEvent property is 1 then the time portion 
        /// of the EventDate MUST be 0 hours UTC.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC02_OperationListItems_fAllDayEvent()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = new DateTime(DateTime.Today.Date.Year, DateTime.Today.Date.Month, DateTime.Today.Date.Day);

            // If fAllDayEvent equal to 1 and the endDate is in same nature day, the end date is equal to 2009-05-19T23:59:59Z
            DateTime endDate = new DateTime(eventDate.Year, eventDate.Month, eventDate.Day, 23, 59, 59);

            // If fAllDayEvent equal to 1, then the eventdate field must be 0 hours UTC as in this example: "2009-05-19T00:00:00Z".
            string timeFormatPattern = @"yyyy-MM-ddT00:00:00Z";
            string eventDateValue = eventDate.ToString(timeFormatPattern);
            string endDateValue = endDate.ToString("yyyy-MM-ddTHH:mm:ssZ");
            string eventTitle = this.GetUniqueListItemTitle("AllDayEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "0" means a single instance.
            recurEventFieldsSetting.Add("EventType", "0");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "1");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            // Set "<ViewFields />" in order to show all fields' value of a list.
            CamlViewFields viewfieds = new CamlViewFields();
            viewfieds.ViewFields = new CamlViewFieldsViewFields();

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");
            string actualEventDateValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventDate");
            string actualEndDateValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EndDate");
            #region Capture code

            DateTime actualEndDate;
            if (!DateTime.TryParse(actualEndDateValue, null, System.Globalization.DateTimeStyles.AdjustToUniversal, out actualEndDate))
            {
                this.Site.Assert.Fail("The EndDate field value should be a valid DateTime format.");
            }

            DateTime actualEventDate;
            if (!DateTime.TryParse(actualEventDateValue, null, System.Globalization.DateTimeStyles.AdjustToUniversal, out actualEventDate))
            {
                this.Site.Assert.Fail("The EventDate field value should be a valid DateTime format.");
            }

            this.Site.Assert.AreEqual<DateTime>(
                               eventDate,
                               actualEventDate,
                               "The EventDate value should equal to the value set in the request of UpdateListItems.");

            // If the EventDate fields' values equal to the values the client set in upon steps and EventDate field in response follow the format "XXXX-XX-XXT00:00:00Z" then capture R862
            this.Site.Assert.AreEqual<int>(
                                        0,
                                        eventDate.Hour,
                                        "The EventDate field of an AllDayEvent type appointment should be a zero hours UTC format.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R862
            this.Site.CaptureRequirement(
                                862,
                                @"[In Appointment-Specific Schema]If the fAllDayEvent property is 1 then the time portion of the EventDate MUST be 0 hours UTC as in this example: ""2009-05-19T00:00:00Z"".");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R300
            this.Site.CaptureRequirementIfAreEqual(
                "0",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType"),
                300,
                "[In Appointment-Specific Schema]If the EventType is something else[0], then RecurrenceID can be empty or missing.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1008
            this.Site.CaptureRequirementIfAreEqual(
                "0",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType"),
                1008,
                "[In EventType][The enumeration value]0[of the type EventType means]Single instance.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC03_OperationListItems_EventTypeRecurring
        /// <summary>
        /// This test case is used to verify if the EventType indicates a recurring event, then fRecurrence MUST be 1.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC03_OperationListItems_EventTypeRecurring()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R277
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                277,
                "[In Appointment-Specific Schema]If the EventType indicates a recurring event, then fRecurrence MUST be 1.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R30001
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType"),
                30001,
                "[In Appointment-Specific Schema]If the EventType is something else[1], then RecurrenceID can be empty or missing.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC04_OperationListItems_EventTypeNotRecurring
        /// <summary>
        /// This test case is used to verify if the EventType does not indicate a recurring event, then fRecurrence MUST be 0.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC04_OperationListItems_EventTypeNotRecurring()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "0" indicates a single instance.
            recurEventFieldsSetting.Add("EventType", "0");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "0");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R278
            this.Site.CaptureRequirementIfAreEqual(
                "0",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                278,
                "[In Appointment-Specific Schema]Otherwise[if the EventType does not indicate a recurring event] fRecurrence MUST be 0.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R279
            this.Site.CaptureRequirementIfIsTrue(
                 "0".Equals(Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"), StringComparison.OrdinalIgnoreCase) || "0".Equals(Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType"), StringComparison.OrdinalIgnoreCase),
                279,
                "[In Appointment-Specific Schema]If EventType indicates a recurring event and fRecurrence is FALSE, then the item is not a recurring event.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC05_OperationListItems_fAllDayEventIsTrue
        /// <summary>
        /// This test case is used to verify the value of fAllDayEvent 1 means true.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC05_OperationListItems_fAllDayEventIsTrue()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "0" indicates a single instance.
            recurEventFieldsSetting.Add("EventType", "0");

            // "1" means this is an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "1");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "0");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            // Set "<ViewFields />" in order to show all fields' value of a list.
            CamlViewFields viewfieds = new CamlViewFields();
            viewfieds.ViewFields = new CamlViewFieldsViewFields();

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R439
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fAllDayEvent"),
                439,
                "[In Appointment-Specific Schema][For fAllDayEvent, the value ]1 means TRUE.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R280
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fAllDayEvent"),
                280,
                "[In Appointment-Specific Schema]fAllDayEvent: A booleanInteger value that specifies whether the appointment is an all-day appointment, as specified in the appointments (section 3.2.1.1) abstract data model.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC06_OperationListItems_fRecurrenceValues1
        /// <summary>
        /// This test case is used to verify the value of fRecurrence 1 means the event is recurring
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC06_OperationListItems_fRecurrenceValues1()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "1" indicates a recurring event. 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If fRecurrence is true, this property MUST contain a valid stringGUID.
            recurEventFieldsSetting.Add("UID", new Guid().ToString());

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R285
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                285,
                "[In Appointment-Specific Schema]fRecurrence: A booleanInteger value that specifies whether the EventType value indicates a recurring event or an exception.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R286
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                286,
                "[In Appointment-Specific Schema][For fRecurrence]1 means it[event] is recurring,");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R984
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                984,
                "[In booleanInteger] The booleanInteger simple type is an integer used to represent a Boolean value.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R985
            this.Site.CaptureRequirementIfAreEqual(
                "1",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                985,
                "[In booleanInteger]If a nonzero value is received in a booleanInteger type field protocol clients and protocol servers MUST treat the nonzero value as 1.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC07_OperationListItems_fRecurrenceValues0
        /// <summary>
        /// This test case is used to verify the value of fRecurrence 0 means the event is not recurring.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC07_OperationListItems_fRecurrenceValues0()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "0" indicates a single instance. 
            recurEventFieldsSetting.Add("EventType", "0");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "0");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R287
            this.Site.CaptureRequirementIfAreEqual(
                "0",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence"),
                287,
                "[In Appointment-Specific Schema][For fRecurrence]0 means it[event] is not[ recurring].");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC08_OperationListItems_ExceptionItemForMasterSeriesItemID
        /// <summary>
        /// This test case is used to verify MasterSeriesItemID exists only for exception items.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC08_OperationListItems_ExceptionItemForMasterSeriesItemID()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item, and its value will be set in RecurrenceID field.
            DateTime overwritedRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");

            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwritedRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            this.Site.Assert.IsTrue(
                            settingsOfException.ContainsKey("MasterSeriesItemID"),
                            "The fields setting collection should contain MasterSeriesItemID field setting.");

            // Set the MasterSeriesItemID field value equal to the recurrence item added in below step.
            string masterSeriesItemIDValue = recurrenceItemId;
            settingsOfException["MasterSeriesItemID"] = masterSeriesItemIDValue;

            this.Site.Assert.IsTrue(
                           settingsOfException.ContainsKey("MasterSeriesItemID"),
                           "The fields setting collection should contain MasterSeriesItemID field setting.");

            this.Site.Assert.IsTrue(
                          settingsOfDailyRecurring.ContainsKey("EventType"),
                          "The fields setting collection should contain EventType field setting.");

            // Set the EventType field value equal to 4, means this item is exception item of a recurrence item.
            string eventTypeValueOfException = "4";
            settingsOfDailyRecurring["EventType"] = eventTypeValueOfException;

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            this.VerifyResponseOfUpdateListItem(updateResultOfException);
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndexOfExceptionItem = this.GetZrowItemIndexByListItemId(zrowItems, exceptionItemIds[0]);
            string actualMasterSeriesItemIDValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfExceptionItem, "ows_MasterSeriesItemID");

            // If the RecurrenceID field value equal to set in the request of UpdateListItems, then capture R293
            this.Site.Assert.AreEqual<string>(
                                    masterSeriesItemIDValue,
                                    actualMasterSeriesItemIDValue,
                                    "The RecurrenceID value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R293
            this.Site.CaptureRequirement(
                           293,
                           @"[In Appointment-Specific Schema]MasterSeriesItemID: This exists only for exception items.");

            string actualEventTypeValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfExceptionItem, "ows_EventType");

            // If the EventType field value equal to set in the request of UpdateListItems(1), then capture R1012
            this.Site.Assert.AreEqual<string>(
                                    eventTypeValueOfException.ToLower(),
                                    actualEventTypeValue.ToLower(),
                                    "The EventType value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1012
            this.Site.CaptureRequirement(
                           1012,
                           @"[In EventType][The enumeration value]4[of the type EventType means]Exception to a recurrence.");
        }
        #endregion

        #region MSOUTSPS_S02_TC09_OperationListItems_ExceptionItemForRecurrenceID
        /// <summary>
        /// This test case is used to verify RecurrenceID is equal to the starting date and time of one instance of a recurrence when the EventType indicates an exception or deleted instance.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC09_OperationListItems_ExceptionItemForRecurrenceID()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting make it begin on 9:00 PM.
            DateTime eventDateOfRecurrence = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 9, 0, 0);

            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);

            this.Site.Assert.IsTrue(
                          settingsOfDailyRecurring.ContainsKey("EventType"),
                          "The fields setting collection should contain EventType field setting.");

            // Set the EventType field value equal to 1, means this item is recurrence item.
            string eventTypeValueOfRecurrence = "1";
            settingsOfDailyRecurring["EventType"] = eventTypeValueOfRecurrence;

            DateTime endDate = eventDateOfRecurrence.AddHours(1);
            settingsOfDailyRecurring["EndDate"] = this.GetGeneralFormatTimeString(endDate);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item, and its value will be set in RecurrenceID field.
            DateTime overwritedRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");

            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwritedRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            this.Site.Assert.IsTrue(
                            settingsOfException.ContainsKey("RecurrenceID"),
                            "The fields setting collection should contain RecurrenceID field setting.");

            // Set the RecurrenceID field value which is point to second recurrence instance.
            string recurrenceIDValue = this.GetGeneralFormatTimeString(overwritedRecerrenceEventDate);
            settingsOfException["RecurrenceID"] = recurrenceIDValue;

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            this.VerifyResponseOfUpdateListItem(updateResultOfException);
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndexOfExceptionItem = this.GetZrowItemIndexByListItemId(zrowItems, exceptionItemIds[0]);
            string actualrecurrenceIDValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfExceptionItem, "ows_RecurrenceID");

            DateTime actualRecurrenceID;
            if (!DateTime.TryParse(actualrecurrenceIDValue, out actualRecurrenceID))
            {
                this.Site.Assert.Fail("The RecurrenceID field value should be valid DateTime format.");
            }

            // If the protocol SUT does return successful repsonse, and in the request the RecurrenceID field value is equal to one of the recurrence instance, and the request is used to create exception recurrence, then capture R299.
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R299
            this.Site.CaptureRequirement(
                           299,
                           @"[In Appointment-Specific Schema]RecurrenceID: RecurrenceID MUST be equal to the starting date and time of one instance of a recurrence when the EventType indicates an exception or deleted instance.");

            int zrowIndexOfRecurrence = this.GetZrowItemIndexByListItemId(zrowItems, recurrenceItemId);
            string actualEventTypeValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfRecurrence, "ows_EventType");

            // If the EventType field value equal to set in the request of UpdateListItems(1), then capture R1099
            this.Site.Assert.AreEqual<string>(
                                    eventTypeValueOfRecurrence.ToLower(),
                                    actualEventTypeValue.ToLower(),
                                    "The EventType value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1009
            this.Site.CaptureRequirement(
                           1009,
                           @"[In EventType][The enumeration value]1[of the type EventType means]Recurring.");
        }
        #endregion

        #region MSOUTSPS_S02_TC10_OperationListItems_UID
        /// <summary>
        /// This test case is used to verify if fRecurrence is true, UID MUST contain a valid stringGUID.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC10_OperationListItems_UID()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            Guid uid = Guid.NewGuid();
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "1" indicates a recurring event. 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If fRecurrence is true, this property MUST contain a valid stringGUID.
            recurEventFieldsSetting.Add("UID", uid.ToString());

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");
            string actualUIDValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_UID");
            #region Capture code

            Guid actualUID;
            if (!Guid.TryParse(actualUIDValue, out actualUID))
            {
                this.Site.Assert.Fail("The UID field should be valid GUID format.");
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R311
            this.Site.CaptureRequirementIfAreEqual(
                uid,
                actualUID,
                311,
                "[In Appointment-Specific Schema]UID: If fRecurrence is true, this property[UID] MUST contain a valid stringGUID (section 2.2.5.11).");

            // Get list items' changes
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // If the UID is existed in list items' changes, and pass the schema validation, than capture R1036, R1037
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItemsOfRecurrenceAppointment, "1");
            string actualUidValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, zrowIndex, "ows_UID");
            this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(actualUidValue),
                                "The UID field should have value.");

            // In current test suite, only the UID field is test able for stringGUID type.
            bool isPassTheSchemaValidation = this.VerifySimpleTypeSchema(zrowItemsOfRecurrenceAppointment);

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1036
            this.Site.CaptureRequirementIfIsTrue(
                isPassTheSchemaValidation,
                1036,
                "[In stringGUID]The stringGUID simple type is a GUID written as a string using hexadecimal digits enclosed by {} and separated by '-',");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1037
            this.Site.CaptureRequirementIfIsTrue(
                isPassTheSchemaValidation,
                1037,
                @"[In stringGUID][The schema definition of stringGUID is:]
                    <s:simpleType name=""stringGUID"">
                      <s:restriction base=""s:string"">
                        <s:maxLength value=""38""/>
                        <s:minLength value=""38""/>
                        <s:pattern value=""\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}"" />
                      </s:restriction>
                    </s:simpleType>");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC11_OperationListItems_XMLTZoneMissing
        /// <summary>
        /// This test case is used to verify if fRecurrence is FALSE, TimeZoneXML can be empty.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC11_OperationListItems_XMLTZoneMissing()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "0" indicates a single instance. 
            recurEventFieldsSetting.Add("EventType", "0");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "0");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R31601
            this.Site.CaptureRequirementIfAreEqual(
                string.Empty,
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_XMLTZone"),
                31601,
                "[In Appointment-Specific Schema]If fRecurrence is FALSE, then this property[TimeZoneXML] [SHOULD be ignored and] can be empty.");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC12_OperationListItems_UIDIgnored
        /// <summary>
        /// This test case is used to verify if fRecurrence is FALSE, UID is ignored.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC12_OperationListItems_UIDIgnored()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region Create a Recurrence AppointMent with invalid UID

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "1" indicates a recurring event. 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // UID is one of the fields used to indicate recurrence changes .
            recurEventFieldsSetting.Add("UID", "invalidUID");

            #endregion

            #region Create a Recurrence AppointMent without UID

            Dictionary<string, string> recurEventFieldsSettingSecond = new Dictionary<string, string>();

            // Setting necessary fields' value
            recurEventFieldsSettingSecond.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSettingSecond.Add("EndDate", endDateValue);

            // "1" indicates a recurring event. 
            recurEventFieldsSettingSecond.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSettingSecond.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSettingSecond.Add("fRecurrence", "1");

            #endregion

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);
            addeditemsOfRecurrence.Add(recurEventFieldsSettingSecond);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.New);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R31301
            this.Site.CaptureRequirementIfAreEqual(
                2,
                updateResult.Results.Length,
                31301,
                "[In Appointment-Specific Schema]If fRecurrence is false, whether the UID is valid or invalid, the server reply the same.");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC13_OperationListItemsForAppointment_XMLTZoneValid
        /// <summary>
        /// This test case is used to verify if EventType is 1, then this property MUST contain a valid TimeZoneXML.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC13_OperationListItemsForAppointment_XMLTZoneValid()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", then this property MUST contain a valid TimeZoneXML.
            recurEventFieldsSetting.Add("XMLTZone", @"<timeZoneRule><standardBias>480</standardBias><additionalDaylightBias>-60</additionalDaylightBias><standardDate><transitionRule month='11' day='su' weekdayOfMonth='first' /><transitionTime>2:0:0</transitionTime></standardDate><daylightDate><transitionRule month='3' day='su' weekdayOfMonth='second' /><transitionTime>2:0:0</transitionTime></daylightDate></timeZoneRule>");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R317
            this.Site.CaptureRequirementIfAreEqual(
                "<timeZoneRule><standardBias>480</standardBias><additionalDaylightBias>-60</additionalDaylightBias><standardDate><transitionRule month='11' day='su' weekdayOfMonth='first' /><transitionTime>2:0:0</transitionTime></standardDate><daylightDate><transitionRule month='3' day='su' weekdayOfMonth='second' /><transitionTime>2:0:0</transitionTime></daylightDate></timeZoneRule>",
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_XMLTZone"),
                317,
                @"[In Appointment-Specific Schema]If EventType is 1 and fAllDayEvent is 1, XMLTZone MUST indicate a time zone with no bias or offset: ""<timeZoneRule><standardBias>0</standardBias><additionalDaylightBias>0</additionalDaylightBias></timeZoneRule>"".");

            #endregion
        }

        #endregion

        #region MSOUTSPS_S02_TC14_RecurrenceAppointmentItem_VerifyUIDField
        /// <summary>
        /// This test case is used to verify UID MUST be changed if the recurrence has been changed or added.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC14_RecurrenceAppointmentItem_VerifyUIDField()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            Guid uid = Guid.NewGuid();

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;
            recurrenceXMLData.recurrence.rule.Item = "7";
            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // "1" indicates a recurring event. 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            List<string> recurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #region Update the recurrence event
            Dictionary<string, string> recurEventFieldsSettingUpdate = new Dictionary<string, string>();
            recurEventFieldsSettingUpdate.Add("ID", recurrenceItemIds[0]);

            // If fRecurrence is true, this property MUST contain a valid stringGUID.
            recurEventFieldsSettingUpdate.Add("UID", uid.ToString());

            List<Dictionary<string, string>> addeditemsOfRecurrenceUpdate = new List<Dictionary<string, string>>();
            addeditemsOfRecurrenceUpdate.Add(recurEventFieldsSettingUpdate);

            List<MethodCmdEnum> cmdsUpdate = new List<MethodCmdEnum>(1);
            cmdsUpdate.Add(MethodCmdEnum.Update);

            // Update the recurrence event
            UpdateListItemsUpdates updatesOfRecurrenceUpdate = this.CreateUpdateListItems(cmdsUpdate, addeditemsOfRecurrenceUpdate, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultUpdate = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrenceUpdate);
            this.VerifyResponseOfUpdateListItem(updateResultUpdate);
            #endregion

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");
            string actualUIDValue = Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_UID");

            #region Capture code

            Guid actualUID;
            if (!Guid.TryParse(actualUIDValue, out actualUID))
            {
                this.Site.Assert.Fail("The UID field should be valid GUID format.");
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R312
            this.Site.CaptureRequirementIfAreEqual<Guid>(
                uid,
                actualUID,
                312,
                "[In Appointment-Specific Schema]UID MUST be changed if, and only if, the recurrence has been changed or added. ");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC15_OperateOnListItems_VerifyChangeTypeValue
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC15_OperateOnListItems_VerifyChangeTypeValue()
        {
            string listId1 = this.AddListToSUT(TemplateType.Generic_List);
            string listId2 = this.AddListToSUT(TemplateType.Generic_List);
            this.AddItemsToList(listId2, 10);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Call GetListItemChangesSinceToken operation to get the changetoken on list 1.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                          listId1,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null);

                if (null == listItemChangesRes || null == listItemChangesRes.listitems || null == listItemChangesRes.listitems.Changes)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken should contain the valid the changes data.");
                }

                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listItemChangesRes.listitems.Changes.LastChangeToken),
                                "The response of GetListItemChangesSinceToken should contain valid ChangeToken.");

                string changeTokenValue = listItemChangesRes.listitems.Changes.LastChangeToken;

                // Add 10 list items into list 2.
                this.AddItemsToList(listId2, 10);

                // Setting "<viewFields />" to view all fields of the list.
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                // Set the changeToken from list 1, it is not valid for list 2, so that the protocol SUT will return ChangeType value in Id element, this means protocol SUT could not query information for list items.
                listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId2,
                                            null,
                                            null,
                                            viewfieds,
                                            null,
                                            null,
                                            changeTokenValue,
                                            null);

                if (null == listItemChangesRes || null == listItemChangesRes.listitems || null == listItemChangesRes.listitems.Changes)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken should contain the valid the changes data.");
                }

                // If the Id element present and the ChangeType value equal to "InvalidToken", then capture 1223
                this.Site.Assert.IsNotNull(
                                        listItemChangesRes.listitems.Changes.Id,
                                        @"The response of GetListItemChangesSinceToken should contain Id element if the request contains invalid changeToken value.");

                this.Site.Assert.IsTrue(
                                       listItemChangesRes.listitems.Changes.Id.ChangeTypeSpecified,
                                       @"The response of GetListItemChangesSinceToken should have changeType value if the request contains invalid changeToken value.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1223
                this.Site.CaptureRequirementIfAreEqual<ChangeTypeEnum>(
                                        ChangeTypeEnum.InvalidToken,
                                        listItemChangesRes.listitems.Changes.Id.ChangeType,
                                        1223,
                                        @"[In GetListItemChangesSinceTokenResponse][If present condition meet][The attribute]Changes.Id.ChangeType is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC16_RecurrenceAppointmentItem_RecurrenceDataMIssing
        /// <summary>
        /// This test case is used to verify if fRecurrence is FALSE, RecurrenceData can be empty or missing.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC16_RecurrenceAppointmentItem_RecurrenceDataMIssing()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "0" means this is not a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "0");

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R29801
            this.Site.CaptureRequirementIfIsTrue(
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_fRecurrence") == "0" && Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_RecurrenceData") == string.Empty,
                29801,
                "[In Appointment-Specific Schema]If fRecurrence is FALSE, then this property[RecurrenceData] [MUST be ignored and ]can be empty or missing.");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC17_RecurrenceAppointmentItem_RecurrenceDataValid
        /// <summary>
        /// This test case is used to verify if EventType is 1, RecurrenceData MUST contain a valid RecurrenceXML.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC17_RecurrenceAppointmentItem_RecurrenceDataValid()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("RecurrenceEvent");

            recurEventFieldsSetting.Add("Title", eventTitle);

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;
            recurrenceXMLData.recurrence.rule.Item = "7";
            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R297
            this.Site.CaptureRequirementIfIsTrue(
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_EventType") == "1" && Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_RecurrenceData") == recurrenceXmlString,
                297,
                "[In Appointment-Specific Schema]RecurrenceData: If EventType is 1, this property MUST contain a valid RecurrenceXML (section 2.2.4.4).");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R898
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                898,
                @"[In Complex Types]The RecurrenceXML complex type contains a RecurrenceDefinition (section 2.2.4.3).
                    <s:complexType name=""RecurrenceXML"">
                      <s:sequence>
                        <s:element name=""recurrence"" type=""s1:RecurrenceDefinition"" />
                        <s:element name=""deleteExceptions"" type=""s:string"" fixed=""true"" minOccurs=""0"" maxOccurs=""1"" />
                      </s:sequence>
                    </s:complexType>");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R877
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                877,
                "[In Complex Types][The complex type]RecurrenceXML Contains a RecurrenceDefinition (section 2.2.4.3).");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R896
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                896,
                @"[In RecurrenceDefinition complex type] The RecurrenceDefinition complex type contains a RecurrenceRule (section 2.2.4.2).   
                    <s:complexType name=""RecurrenceDefinition"">
                      <s:sequence>
                        <s:element name=""rule"" type=""s1:RecurrenceRule"" />
                      </s:sequence>
                    </s:complexType>");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R876
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                876,
                "[In Complex Types][The complex type]RecurrenceDefinition Contains a RecurrenceRule (section 2.2.4.2).");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R888
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                888,
                @"[In RecurrenceRule complex type] The RecurrenceRule complex type defines when a recurrence takes place.
                    <s:complexType name=""RecurrenceRule"">
                      <s:sequence>
                        <s:element name=""firstDayOfWeek"" type=""s1:DayOfWeekOrMonth"" />
                        <s:element name=""repeat"" type=""s1:RepeatPattern"" />
                        <s:choice>
                          <s:element name=""windowEnd"" type=""s:dateTime"" />
                          <s:element name=""repeatForever"">
                            <s:simpleType>
                              <s:restriction base=""s:string"">            
                                <s:enumeration value=""FALSE"" />
                              </s:restriction>
                            </s:simpleType>
                          </s:element>
                          <s:element name=""repeatInstances"" type=""s:integer"" />
                        </s:choice>
                      </s:sequence>
                    </s:complexType>");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R875
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                875,
                "[In Complex Types][The complex type]RecurrenceRule Defines when a recurrence takes place.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R878
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                878,
                "[In Complex Types][The complex type]RepeatPattern Contains a choice of elements which describe what days a recurrence occurs on.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R902
            this.Site.CaptureRequirementIfIsTrue(
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML)),
                902,
                @"[In Complex Types]The RepeatPattern complex type contains a choice of elements which describe what days a recurrence occurs on.
                    <s:complexType name=""RepeatPattern"">
                      <s:choice>
                        <s:element name=""daily"">
                          <s:complexType>
                            <s:simpleContent>
                              <s:extension base=""s:string"">
                                <s:attribute name=""su"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""mo"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""tu"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""we"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""th"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""fr"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""sa"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""weekFrequency"" type=""s:integer"" default=""1"" use=""optional"" />
                              </s:extension>
                            </s:simpleContent>
                          </s:complexType>
                        </s:element>
                        <s:element name=""monthlyByDay"">
                          <s:complexType>
                            <s:simpleContent>
                              <s:extension base=""s:string"">
                                <s:attribute name=""su"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""mo"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""tu"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""we"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""th"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""fr"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""sa"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""day"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""weekday"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""weekend_day"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""monthFrequency"" type=""s:integer"" default=""1"" use=""optional"" />
                                <s:attribute name=""weekdayOfMonth"" type=""s1:WeekdayOfMonth"" default=""first"" use=""optional"" />
                              </s:extension>
                            </s:simpleContent>
                          </s:complexType>
                        </s:element>
                        <s:element name=""monthly"">
                          <s:complexType>
                            <s:simpleContent>
                              <s:extension base=""s:string"">
                                <s:attribute name=""monthFrequency"" type=""s:integer"" default=""1"" use=""optional"" />
                               <s:attribute name=""day"" type=""s:integer"" default=""1"" use=""optional"" />
                              </s:extension>
                            </s:simpleContent>
                          </s:complexType>
                        </s:element>
                        <s:element name=""yearly"">
                          <s:complexType>
                            <s:simpleContent>
                              <s:extension base=""s:string"">
                                <s:attribute name=""yearFrequency"" type=""s:integer"" default=""1"" use=""optional"" />
                                <s:attribute name=""month"" type=""s:integer"" default=""1"" use=""optional"" />
                                <s:attribute name=""day"" type=""s:integer"" default=""1"" use=""optional"" />
                              </s:extension>
                           </s:simpleContent>
                          </s:complexType>
                        </s:element>
                        <s:element name=""yearlyByDay"">
                          <s:complexType>
                            <s:simpleContent>
                              <s:extension base=""s:string"">
                                <s:attribute name=""su"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""mo"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                               <s:attribute name=""tu"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                              <s:a name=""we"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""th"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""fr"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""sa"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""day"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""weekday"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""weekend_day"" type=""s1:TrueFalseDOW"" default=""FALSE"" use=""optional"" />
                                <s:attribute name=""yearFrequency"" type=""s:integer"" default=""1"" use=""optional"" />
                                <s:attribute name=""month"" type=""s:integer"" default=""1"" use=""optional"" />
                                <s:attribute name=""weekdayOfMonth"" type=""s1:WeekdayOfMonth"" default=""first"" use=""optional"" />
                             </s:extension>
                            </s:simpleContent>
                          </s:complexType>
                        </s:element>
                      </s:choice>
                    </s:complexType>
                    ");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC18_RecurrenceAppointmentItem_RecurrenceRulewindowEnd
        /// <summary>
        /// This test case is used to verify RecurrenceRule complex type and windowEnd element.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC18_RecurrenceAppointmentItem_RecurrenceRulewindowEnd()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);
            DateTime windowEnd = endDate.Date.AddHours(-1);

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Setting the windowEnd.
            recurrenceXMLData.recurrence.rule.Item = windowEnd;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;

            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1262
            this.Site.CaptureRequirementIfAreEqual(
                recurrenceXmlString,
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_RecurrenceData"),
                1262,
                "[In RecurrenceRule complex type] windowEnd: Client set this value[windowEnd], and later client retrieve the value[windowEnd] from the server, they should be equal.");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC19_RecurrenceAppointmentItem_RecurrenceRulerepeatForever
        /// <summary>
        /// This test case is used to verify RecurrenceRule complex type and repeatForever element.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC19_RecurrenceAppointmentItem_RecurrenceRulerepeatForever()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            Dictionary<string, string> recurEventFieldsSetting = new Dictionary<string, string>();
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.Date.AddHours(1);
            string eventDateValue = this.GetGeneralFormatTimeString(eventDate);
            string endDateValue = this.GetGeneralFormatTimeString(endDate);

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;
            recurrenceXMLData.recurrence.rule.Item = "FALSE";

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();

            // Setting the dayFrequencyValue
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;

            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting necessary fields' value
            recurEventFieldsSetting.Add("EventDate", eventDateValue);

            // The ending date and time of the appointment.
            recurEventFieldsSetting.Add("EndDate", endDateValue);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting.Add("EventType", "1");

            // "0" means this is not an all-day event.
            recurEventFieldsSetting.Add("fAllDayEvent", "0");

            // "1" means this is a recurrence event.
            recurEventFieldsSetting.Add("fRecurrence", "1");

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting.Add("RecurrenceData", recurrenceXmlString);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a  recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1263
            this.Site.CaptureRequirementIfAreEqual(
                recurrenceXmlString,
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_RecurrenceData"),
                1263,
                "[In RecurrenceRule complex type] repeatForever: Client set this value[repeatForever], and later client retrieve the value[repeatForever] from the server, they should be equal.");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC20_RecurrenceAppointmentItem_RecurrenceRulerepeatInstances
        /// <summary>
        /// This test case is used to verify RecurrenceRule complex type and repeatInstances element.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC20_RecurrenceAppointmentItem_RecurrenceRulerepeatInstances()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            DateTime eventDate = DateTime.Today.Date.AddDays(1);

            Dictionary<string, string> recurEventFieldsSetting = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("YearlyRecurrenceItem"), eventDate, "1", 10);
            this.Site.Assert.IsTrue(
                            recurEventFieldsSetting.ContainsKey("RecurrenceData"),
                            "The fields setting collection should contain RecurrenceData field setting.");

            // Recurring data
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();
            RepeatPatternYearlyByDay repeatPatternYearlyByDay = new RepeatPatternYearlyByDay();

            // Setting the repeatPatternYearlyByDay value
            repeatPatternYearlyByDay.mo = TrueFalseDOW.TRUE;
            repeatPatternYearlyByDay.moSpecified = true;

            // Each year occurs one time.
            repeatPatternYearlyByDay.yearFrequency = "1";

            // Occurs on second week
            repeatPatternYearlyByDay.weekdayOfMonth = WeekdayOfMonth.second;
            repeatPatternYearlyByDay.weekdayOfMonthSpecified = true;

            // Occurs on current month
            repeatPatternYearlyByDay.month = eventDate.Month.ToString();
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternYearlyByDay;

            // Setting the repeatInstances to repeat ten times.
            recurrenceXMLData.recurrence.rule.Item = "10";

            string recurrenceXmlString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // If the EventType indicates a recurring event, then fRecurrence MUST be "1". 
            recurEventFieldsSetting["EventType"] = "1";

            // "0" means this is not an all-day event.
            recurEventFieldsSetting["fAllDayEvent"] = "0";

            // "1" means this is a recurrence event.
            recurEventFieldsSetting["fRecurrence"] = "1";

            // If EventType is "1", this property MUST contain a valid RecurrenceXML.
            recurEventFieldsSetting["RecurrenceData"] = recurrenceXmlString;
            DateTime endTime = eventDate.AddHours(1);

            // For yearly recurrence item, the end date must larger than the current year + repeatInstances.
            recurEventFieldsSetting["EndDate"] = this.GetGeneralFormatTimeString(endTime.AddYears(100));
            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(recurEventFieldsSetting);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // add a recurrence appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowitems = this.GetListItemsChangesFromSUT(listId);
            int listItemIndex = this.GetZrowItemIndexByListItemId(zrowitems, "1");

            #region Capture code

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1266
            this.Site.CaptureRequirementIfAreEqual(
                recurrenceXmlString,
                Common.GetZrowAttributeValue(zrowitems, listItemIndex, "ows_RecurrenceData"),
                1266,
                "[In RecurrenceRule complex type] repeatInstances: Client set this value[repeatInstances], and later client retrieve the value[repeatInstances] from the server, they should be equal.specified by this element.");

            #endregion
        }
        #endregion

        #region MSOUTSPS_S02_TC21_RecurrenceAppointmentItem_RecurrenceDefinition
        /// <summary>
        /// This test case is used to verify RecurrenceDefinition complex type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC21_RecurrenceAppointmentItem_RecurrenceDefinition()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("YearlyRecurrenceItem"), eventDateOfRecurrence, "1", 10);
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                            "The fields setting collection should contain RecurrenceData field setting.");

            // Setting a yearly recurrence setting for RecurrenceData field.
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Set the repeatInstances value, it will end after 3 occurrences.
            int repeatInstanceValue = 3;
            recurrenceXMLData.recurrence.rule.Item = repeatInstanceValue.ToString();

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the RepeatPatternYearly, make the recurrence occur on 04-01 for every year.
            RepeatPatternYearly repeatPatternYearlyData = new RepeatPatternYearly();
            repeatPatternYearlyData.day = "1";
            repeatPatternYearlyData.month = "4";
            repeatPatternYearlyData.yearFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternYearlyData;
            string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting the RecurrenceData field with RepeatPattern
            settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

            // Setting the EndDate field to match the Yearly and repeatInstances setting.
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("EndDate"),
                            "The fields setting collection should contain EndDate field setting.");

            // The recurrence item end time.
            DateTime endTime = eventDateOfRecurrence.AddHours(1);

            // For yearly recurrence item, the end date must larger than the current year + repeatInstances.
            settingsOfDailyRecurring["EndDate"] = this.GetGeneralFormatTimeString(endTime.AddYears(100));

            #region add a recurrence item

            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

            // Add current list item
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #endregion add a recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualRecurrenceXMLStringValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_RecurrenceData");

            // If the RecurrenceData field value equal to set in the request of UpdateListItems, that means setting of RecurrenceDefinition.rule is set in the protocol SUT. If this verification passes then capture R897
            this.Site.Assert.AreEqual<string>(
                                    recurrenceXMLString.ToLower(),
                                    actualRecurrenceXMLStringValue.ToLower(),
                                    "The RecurrenceData value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R897
            this.Site.CaptureRequirement(
                           897,
                           @"[In RecurrenceDefinition complex type]rule: Contains a recurrence rule that defines a recurrence.");
        }

        #endregion

        #region MSOUTSPS_S02_TC22_RecurrenceAppointmentItem_RepeatPattern
        /// <summary>
        /// This test case is used to verify RepeatPattern complex type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC22_RecurrenceAppointmentItem_RepeatPattern()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                            "The fields setting collection should contain RecurrenceData field setting.");

            // Setting a daily recurrence setting for RecurrenceData field.
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.su;

            // Set the windowEnd value, the recurrence will be end on one week from current date.
            DateTime winEndDate = eventDateOfRecurrence.Date.AddDays(7);
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the RepeatPatternDaily
            RepeatPatternDaily repeatPatternDailyData = new RepeatPatternDaily();
            repeatPatternDailyData.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDailyData;
            string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);

            // Setting the RecurrenceData field with RepeatPattern.daily
            settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

            #region add a recurrence item

            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

            // Add current list item
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #endregion add a recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualRecurrenceXMLStringValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_RecurrenceData");

            // If the RecurrenceData field value equal to set in the request of UpdateListItems, that means setting of RepeatPattern.daily is set in the protocol SUT. If this verification passes then capture R903
            this.Site.Assert.AreEqual<string>(
                                    recurrenceXMLString.ToLower(),
                                    actualRecurrenceXMLStringValue.ToLower(),
                                    "The RecurrenceData value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R903
            this.Site.CaptureRequirement(
                           903,
                           @"[In RepeatPattern complex type]daily: If the recurrence is a daily recurrence, the daily element will be present.");
        }

        #endregion

        #region MSOUTSPS_S02_TC23_OperationListItemsForAppointment_TransitionDate
        /// <summary>
        /// This test case is used to verify TransitionDate complex type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC23_OperationListItemsForAppointment_TransitionDate()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // Get a Pacific time zone setting.
            TimeZoneXML timeZoneXml = this.GetCustomPacificTimeZoneXmlSetting();

            // The timeZoneXml.timeZoneRule.standardDate and timeZoneXml.timeZoneRule.daylightDate element is the type of TransitionDate.
            // Setting TransitionDate.transitionRule.dayOfMonth and TransitionDate.transitionTime
            timeZoneXml.timeZoneRule.daylightDate.transitionRule.dayOfMonth = "12";
            timeZoneXml.timeZoneRule.standardDate.transitionTime = "3:0:0";
            string timeZoneXmlString = this.GetTimeZoneXMLString(timeZoneXml);

            #region add a recurrence item

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("XMLTZone"),
                            "The fields setting collection should contain XMLTZone field setting.");

            // Setting the XMLTZone field value with setting of TransitionDate.transitionRule.dayOfMonth and TransitionDate.transitionTime
            settingsOfDailyRecurring["XMLTZone"] = timeZoneXmlString;

            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

            // Add current list item
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #endregion add a recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualXMLTZoneValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_XMLTZone");

            // If the XMLTZone field value equal to set in the request of UpdateListItems, that means setting of TransitionDate.transitionRule.dayOfMonth and TransitionDate.transitionTime are set in the protocol SUT. If this verification passes then capture R1268, R1269
            this.Site.Assert.AreEqual<string>(
                                    timeZoneXmlString.ToLower(),
                                    actualXMLTZoneValue.ToLower(),
                                    "The XMLTZone value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1268
            this.Site.CaptureRequirement(
                           1268,
                           @"[In TransitionDate complex type] transitionRule.dayOfMonth: Client set this value[transitionRule.dayOfMonth], and later client retrieve the value[transitionRule.dayOfMonth] from the server, they should be equal.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1269
            this.Site.CaptureRequirement(
                           1269,
                           @"[In TransitionDate complex type] transitionTime: Client set this value[transitionTime], and later client retrieve the value[transitionTime] from the server, they should be equal.");
        }

        #endregion

        #region MSOUTSPS_S02_TC24_OperationListItems_TimeZoneRule
        /// <summary>
        /// This test case is used to verify TimeZoneRule complex type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC24_OperationListItems_TimeZoneRule()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // Get a Pacific time zone setting, it set the standardBias, additionalDaylightBias, standardDate and daylightDate setting with correct Pacific time zone.
            TimeZoneXML timeZoneXml = this.GetCustomPacificTimeZoneXmlSetting();
            string timeZoneXmlString = this.GetTimeZoneXMLString(timeZoneXml);

            #region add a recurrence item

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("XMLTZone"),
                            "The fields setting collection should contain XMLTZone field setting.");

            // Setting the XMLTZone field value.
            settingsOfDailyRecurring["XMLTZone"] = timeZoneXmlString;

            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

            // Add current list item
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #endregion add a recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualXMLTZoneValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_XMLTZone");

            // If the XMLTZone value equal to set in the request of UpdateListItems, then capture R946, R947, R948, R949
            this.Site.Assert.AreEqual<string>(
                                    timeZoneXmlString.ToLower(),
                                    actualXMLTZoneValue.ToLower(),
                                    "The XMLTZone value equal to set in the request of UpdateListItems");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R946
            this.Site.CaptureRequirement(
                           946,
                           @"[In TimeZoneRule complex type]standardBias: An integer that specifies the time difference, in minutes, from Coordinated Universal Time (UTC).");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R947
            this.Site.CaptureRequirement(
                           947,
                           @"[In TimeZoneRule complex type]additionalDaylightBias: An integer that specifies in minutes the time that is added to the standardBias while the time zone is between daylightDate and standardDate.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R948
            this.Site.CaptureRequirement(
                           948,
                           @"[In TimeZoneRule complex type]standardDate: The date and time after which only standardBias is applied.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R949
            this.Site.CaptureRequirement(
                           949,
                           @"[In TimeZoneRule complex type]daylightDate: The date and time after which standardBias plus additionalDaylightBias is applied.");
        }

        #endregion

        #region MSOUTSPS_S02_TC25_OperationListItems_DayOfWeekSimpleType
        /// <summary>
        /// This test case is used to verify DayOfWeek simple type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC25_OperationListItems_DayOfWeekSimpleType()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            TimeZoneXML timeZoneXml = this.GetCustomPacificTimeZoneXmlSetting();

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
            Dictionary<string, string> listItemTimeZoneSettings = new Dictionary<string, string>();

            #region add 7 recurrence items for each value of DayOfWeek enum

            // Add recurrence items for each value of DayOfWeek enum.
            string[] namesOfDayOfDayOfWeek = Enum.GetNames(typeof(DayOfWeek));
            foreach (string enumNameItem in namesOfDayOfDayOfWeek)
            {
                List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
                List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

                string eventTitle = this.GetUniqueListItemTitle("DayOfWeekInTransitionRule");

                this.Site.Assert.IsTrue(
                               settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                               "The fields setting should contain RecurrenceData field's setting.");

                this.Site.Assert.IsTrue(
                                     settingsOfDailyRecurring.ContainsKey("Title"),
                                     "The fields setting should contain Title field's setting.");

                settingsOfDailyRecurring["Title"] = eventTitle;

                // Update the TimeZoneXML field, set the timeZoneXml.timeZoneRule.standardDate.transitionRule.day
                DayOfWeek currentDayOfWeekOrMonthValue = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), enumNameItem, true);
                timeZoneXml.timeZoneRule.standardDate.transitionRule.day = currentDayOfWeekOrMonthValue;
                string timeZoneXmlString = this.GetTimeZoneXMLString(timeZoneXml);
                settingsOfDailyRecurring["XMLTZone"] = timeZoneXmlString;

                // Add current list item
                addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
                cmds.Add(MethodCmdEnum.New);

                UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

                // Get the listItem id and the related RecurrenceData
                this.VerifyResponseOfUpdateListItem(updateResult);
                List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
                listItemTimeZoneSettings.Add(addedRecurrenceItemIds[0], timeZoneXmlString);
            }

            this.Site.Assert.AreEqual<int>(
                                    namesOfDayOfDayOfWeek.Length,
                                    listItemTimeZoneSettings.Count,
                                    "There should be match number of list items' TimeZone data setting");

            #endregion add 7 recurrence items for each value of DayOfWeek enum

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            this.Site.Assert.AreEqual<int>(
                                    namesOfDayOfDayOfWeek.Length,
                                    zrowItems.Count(),
                                    "The current list should contain [{0}] list items.",
                                    namesOfDayOfDayOfWeek.Length);

            // Verify each WeekdayOfMonth value can be set in protocol SUT successfully.
            foreach (KeyValuePair<string, string> listItemRecurrenceDataItem in listItemTimeZoneSettings)
            {
                string currentListItemId = listItemRecurrenceDataItem.Key;
                int zrowIndexOfCurrentListItemId = this.GetZrowItemIndexByListItemId(zrowItems, currentListItemId);
                string actualRecurrenceData = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfCurrentListItemId, "ows_XMLTZone");
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(actualRecurrenceData),
                                "The actual TimeZone field should have value.");

                string expectedRecurrenceData = listItemRecurrenceDataItem.Value.ToLower();
                this.Site.Assert.AreEqual<string>(
                                    expectedRecurrenceData.ToLower(),
                                    actualRecurrenceData.ToLower(),
                                    "The TimeZone field value should match the value set in the request of UpdateListItems operation.");
            }

            // If upon verification pass, then capture R994, R995, R996, R997, R998, R999, R1000 
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R994
            this.Site.CaptureRequirement(
                           994,
                           @"[In DayOfWeek][The enumeration value]su[of the type DayOfWeek means]Sunday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R995
            this.Site.CaptureRequirement(
                           995,
                           @"[In DayOfWeek][The enumeration value]mo[of the type DayOfWeek means]Monday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R996
            this.Site.CaptureRequirement(
                           996,
                           @"[In DayOfWeek][The enumeration value]tu[of the type DayOfWeek means]Tuesday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R997
            this.Site.CaptureRequirement(
                           997,
                           @"[In DayOfWeek][The enumeration value]we[of the type DayOfWeek means]Wednesday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R998
            this.Site.CaptureRequirement(
                           998,
                           @"[In DayOfWeek][The enumeration value]th[of the type DayOfWeek means]Thursday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R999
            this.Site.CaptureRequirement(
                           999,
                           @"[In DayOfWeek][The enumeration value]fr[of the type DayOfWeek means]Friday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1000
            this.Site.CaptureRequirement(
                           1000,
                           @"[In DayOfWeek][The enumeration value]sa[of the type DayOfWeek means]Saturday");
        }

        #endregion

        #region MSOUTSPS_S02_TC26_OperationListItems_DayOfWeekOrMonthSimpleType
        /// <summary>
        /// This test case is used to verify DayOfWeekOrMonth simple type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC26_OperationListItems_DayOfWeekOrMonthSimpleType()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add 5 recurrence item

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);

            // Setting common recurrenceXMLData setting
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();

            // Set the windowEnd value, the recurrence will be end on one week later, and the repeat frequency is daily.
            DateTime winEndDate = eventDateOfRecurrence.Date.AddDays(7);
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern, set it as weekly repeat.
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the monthFrequency. every month repeat.
            RepeatPatternDaily repeatPatternDaily = new RepeatPatternDaily();
            repeatPatternDaily.dayFrequency = "1";
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternDaily;

            // Add recurrence items with 10 values of WeekdayOfMonth enum.
            string[] namesOfDayOfWeekOrMonth = Enum.GetNames(typeof(DayOfWeekOrMonth));
            Dictionary<string, string> listItemRecurrenceDataSettings = new Dictionary<string, string>();
            foreach (string enumNameItem in namesOfDayOfWeekOrMonth)
            {
                List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
                List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

                string eventTitle = this.GetUniqueListItemTitle("DayOfWeekOrMonth");
                Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
                this.Site.Assert.IsTrue(
                               settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                               "The fields setting should contain RecurrenceData field's setting.");

                this.Site.Assert.IsTrue(
                                     settingsOfDailyRecurring.ContainsKey("Title"),
                                     "The fields setting should contain Title field's setting.");

                settingsOfDailyRecurring["Title"] = eventTitle;

                // Update the RecurrenData field with recurrence.rule.firstDayOfWeek setting
                DayOfWeekOrMonth currentDayOfWeekOrMonthValue = (DayOfWeekOrMonth)Enum.Parse(typeof(DayOfWeekOrMonth), enumNameItem, true);
                recurrenceXMLData.recurrence.rule.firstDayOfWeek = currentDayOfWeekOrMonthValue;

                string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);
                settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

                // Add current list item
                addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
                cmds.Add(MethodCmdEnum.New);

                UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

                // Get the listItem id and the related RecurrenceData
                this.VerifyResponseOfUpdateListItem(updateResult);
                List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
                listItemRecurrenceDataSettings.Add(addedRecurrenceItemIds[0], recurrenceXMLString);
            }

            this.Site.Assert.AreEqual<int>(
                                    namesOfDayOfWeekOrMonth.Length,
                                    listItemRecurrenceDataSettings.Count,
                                    "There should be match number of list items' Recurrence data setting");

            #endregion  Add 10 recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            this.Site.Assert.AreEqual<int>(
                                    namesOfDayOfWeekOrMonth.Length,
                                    zrowItems.Count(),
                                    "The current list should contain [{0}] list items.",
                                    namesOfDayOfWeekOrMonth.Length);

            // Verify each WeekdayOfMonth value can be set in protocol SUT successfully.
            foreach (KeyValuePair<string, string> listItemRecurrenceDataItem in listItemRecurrenceDataSettings)
            {
                string currentListItemId = listItemRecurrenceDataItem.Key;
                int zrowIndexOfCurrentListItemId = this.GetZrowItemIndexByListItemId(zrowItems, currentListItemId);
                string actualRecurrenceData = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfCurrentListItemId, "ows_RecurrenceData");
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(actualRecurrenceData),
                                "The actual RecurrenceData field should have value.");

                string expectedRecurrenceData = listItemRecurrenceDataItem.Value.ToLower();
                this.Site.Assert.AreEqual<string>(
                                    expectedRecurrenceData.ToLower(),
                                    actualRecurrenceData.ToLower(),
                                    "The RecurrenceData field value should match the value set in the request of UpdateListItems operation.");
            }

            // If upon verification pass, then capture R1001, R10031, R10032, R10033, R10034, R10035, R10036, R10037, R1004, R1005, R1006
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1001
            this.Site.CaptureRequirement(
                            1001,
                            @"[In DayOfWeekOrMonth] The DayOfWeekOrMonth simple type specifies a day of the week or a day of the month.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1004
            this.Site.CaptureRequirement(
                            1004,
                            @"[In DayOfWeekOrMonth][The enumeration value]day[of the type DayOfWeekOrMonth means]Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, and Saturday are allowed.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1005
            this.Site.CaptureRequirement(
                            1005,
                            @"[In DayOfWeekOrMonth][The enumeration value]weekday[of the type DayOfWeekOrMonth means]Monday, Tuesday, Wednesday, Thursday, and Friday are allowed.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1006
            this.Site.CaptureRequirement(
                            1006,
                            @"[In DayOfWeekOrMonth][The enumeration value]weekend_day[of the type DayOfWeekOrMonth means]Sunday and Saturday are allowed.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10031
            this.Site.CaptureRequirement(
                            10031,
                            @"[In DayOfWeekOrMonth][The enumeration value]su[of the type DayOfWeekOrMonth means]Sunday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10032
            this.Site.CaptureRequirement(
                            10032,
                            @"[In DayOfWeekOrMonth][The enumeration value]mo[of the type DayOfWeekOrMonth means]Monday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10033
            this.Site.CaptureRequirement(
                            10033,
                            @"[In DayOfWeekOrMonth][The enumeration value]tu[of the type DayOfWeekOrMonth means]Tuesday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10034
            this.Site.CaptureRequirement(
                            10034,
                            @"[In DayOfWeekOrMonth][The enumeration value]we[of the type DayOfWeekOrMonth means]Wednesday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10035
            this.Site.CaptureRequirement(
                            10035,
                            @"[In DayOfWeekOrMonth][The enumeration value]th[of the type DayOfWeekOrMonth means]Thursday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10036
            this.Site.CaptureRequirement(
                            10036,
                            @"[In DayOfWeekOrMonth][The enumeration value]fr[of the type DayOfWeekOrMonth means]Friday");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10037
            this.Site.CaptureRequirement(
                            10037,
                            @"[In DayOfWeekOrMonth][The enumeration value]sa[of the type DayOfWeekOrMonth means]Saturday");
        }

        #endregion

        #region MSOUTSPS_S02_TC27_RecurrenceAppointmentItem_TrueFalseDOWSimleType
        /// <summary>
        /// This test case is used to verify TrueFalseDOW simple type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC27_RecurrenceAppointmentItem_TrueFalseDOWSimleType()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("WeeklyRecurrence"), eventDateOfRecurrence, "1", 10);

            // Setting common recurrenceXMLData setting
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.mo;

            // Set the windowEnd value, the recurrence will be end on one month later, and the repeat frequency is weekly.
            DateTime winEndDate = eventDateOfRecurrence.Date.AddMonths(1);
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern, set it as weekly repeat.
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the monthFrequency. every month repeat.
            RepeatPatternWeekly repeatPatternWeekly = new RepeatPatternWeekly();
            repeatPatternWeekly.weekFrequency = "1";

            // Set TrueFalseDOW.True
            repeatPatternWeekly.mo = TrueFalseDOW.TRUE;
            repeatPatternWeekly.moSpecified = true;

            // Set TrueFalseDOW.FALSE
            repeatPatternWeekly.tu = TrueFalseDOW.FALSE;
            repeatPatternWeekly.tuSpecified = true;

            // Set TrueFalseDOW.@true
            repeatPatternWeekly.we = TrueFalseDOW.@true;
            repeatPatternWeekly.weSpecified = true;

            // Set TrueFalseDOW.@false
            repeatPatternWeekly.th = TrueFalseDOW.@false;
            repeatPatternWeekly.thSpecified = true;

            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternWeekly;
            string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);
            this.Site.Assert.IsTrue(
                            settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                            "The fields setting should contain RecurrenceData field setting.");

            settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();

            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
            cmds.Add(MethodCmdEnum.New);

            // Add list item with specified TrueFalseDOW values
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id and the related RecurrenceData
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndexOfAddedItem = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualRecurrenceData = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfAddedItem, "ows_RecurrenceData");

            // If the actual RecurrenceData field value is equal to the value set in the request, then capture R1041, R1042, R1043, R1044
            this.Site.Assert.AreEqual<string>(
                                recurrenceXMLString.ToLower(),
                                actualRecurrenceData.ToLower(),
                                "The actual RecurrenceData field value is equal to the value set in the request of UpdateListItems operation.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1041
            this.Site.CaptureRequirement(
                        1041,
                        @"[TrueFalseDOW][The enumeration value]TRUE[of the type TrueFalseDOW means]true");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1042
            this.Site.CaptureRequirement(
                        1042,
                        @"[TrueFalseDOW][The enumeration value]FALSE[of the type TrueFalseDOW means]false");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1043
            this.Site.CaptureRequirement(
                        1043,
                        @"[TrueFalseDOW][The enumeration value]true[of the type TrueFalseDOW means]true");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1044
            this.Site.CaptureRequirement(
                        1044,
                        @"[TrueFalseDOW][The enumeration value]false[of the type TrueFalseDOW means]false");
        }

        #endregion

        #region MSOUTSPS_S02_TC28_OperationListItems_WeekdayOfMonth
        /// <summary>
        /// This test case is used to verify WeekdayOfMonth simple type.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC28_OperationListItems_WeekdayOfMonth()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add 5 recurrence item

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);

            // Setting common recurrenceXMLData setting
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();
            recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.mo;

            // Set the windowEnd value, the recurrence will be end on one year later, and the repeat frequency is month.
            DateTime winEndDate = eventDateOfRecurrence.Date.AddYears(1);
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern, set it as weekly repeat.
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the monthFrequency. every month repeat.
            RepeatPatternMonthlyByDay repeatPatternMonthlyByDay = new RepeatPatternMonthlyByDay();
            repeatPatternMonthlyByDay.monthFrequency = 1;
            repeatPatternMonthlyByDay.monthFrequencySpecified = true;
            repeatPatternMonthlyByDay.fr = TrueFalseDOW.TRUE;
            repeatPatternMonthlyByDay.moSpecified = true;
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternMonthlyByDay;

            // Add 5 recurrence items with 5 values of WeekdayOfMonth enum.
            string[] namesOfWeekdayOfMonthEnum = Enum.GetNames(typeof(WeekdayOfMonth));
            Dictionary<string, string> listItemRecurrenceDataSettings = new Dictionary<string, string>();
            foreach (string enumNameItem in namesOfWeekdayOfMonthEnum)
            {
                List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
                List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);

                string eventTitle = this.GetUniqueListItemTitle("MonthlyByDay");
                Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
                this.Site.Assert.IsTrue(
                               settingsOfDailyRecurring.ContainsKey("RecurrenceData"),
                               "The fields setting should contain RecurrenceData field's setting.");

                this.Site.Assert.IsTrue(
                                     settingsOfDailyRecurring.ContainsKey("Title"),
                                     "The fields setting should contain Title field's setting.");

                settingsOfDailyRecurring["Title"] = eventTitle;

                RepeatPatternMonthlyByDay currentRepeatPattern = recurrenceXMLData.recurrence.rule.repeat.Item as RepeatPatternMonthlyByDay;
                this.Site.Assert.IsNotNull(currentRepeatPattern, "The RepeatPattern should be RepeatPatternMonthlyByDay type.");

                // Update the RecurrenData field
                WeekdayOfMonth currentWeekdayOfMonthValue = (WeekdayOfMonth)Enum.Parse(typeof(WeekdayOfMonth), enumNameItem, true);
                currentRepeatPattern.weekdayOfMonthSpecified = true;
                currentRepeatPattern.weekdayOfMonth = currentWeekdayOfMonthValue;
                string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);
                settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

                // Add current list item
                addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
                cmds.Add(MethodCmdEnum.New);

                UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

                // Get the listItem id and the related RecurrenceData
                this.VerifyResponseOfUpdateListItem(updateResult);
                List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
                listItemRecurrenceDataSettings.Add(addedRecurrenceItemIds[0], recurrenceXMLString);
            }

            this.Site.Assert.AreEqual<int>(
                                    namesOfWeekdayOfMonthEnum.Length,
                                    listItemRecurrenceDataSettings.Count,
                                    "There should be match number of list items' Recurrence data setting");

            #endregion  Add 5 recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            this.Site.Assert.AreEqual<int>(
                                    namesOfWeekdayOfMonthEnum.Length,
                                    zrowItems.Count(),
                                    "The current list should contain [{0}] list items.",
                                    namesOfWeekdayOfMonthEnum.Length);

            // Verify each WeekdayOfMonth value can be set in protocol SUT successfully.
            foreach (KeyValuePair<string, string> listItemRecurrenceDataItem in listItemRecurrenceDataSettings)
            {
                string currentListItemId = listItemRecurrenceDataItem.Key;
                int zrowIndexOfCurrentListItemId = this.GetZrowItemIndexByListItemId(zrowItems, currentListItemId);
                string actualRecurrenceData = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfCurrentListItemId, "ows_RecurrenceData");
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(actualRecurrenceData),
                                "The actual RecurrenceData field should have value.");

                string expectedRecurrenceData = listItemRecurrenceDataItem.Value.ToLower();
                this.Site.Assert.AreEqual<string>(
                                    expectedRecurrenceData.ToLower(),
                                    actualRecurrenceData.ToLower(),
                                    "The RecurrenceData field value should match the value set in the request of UpdateListItems operation.");
            }

            // If passes upon verification, then capture R1045, R1047, R1048, R1049, R1050, R1051
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1045
            this.Site.CaptureRequirement(
                                1045,
                                @"[In WeekdayOfMonth] When combined with a DayOfWeek or DayOfWeekOrMonth value, the WeekdayOfMonth simple type specifies a day of a week or month.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1047
            this.Site.CaptureRequirement(
                                1047,
                                @"[In WeekdayOfMonth][The enumeration value]first[of the type WeekdayOfMonth means]First");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1048
            this.Site.CaptureRequirement(
                                1048,
                                @"[In WeekdayOfMonth][The enumeration value]second[of the type WeekdayOfMonth means]Second");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1049
            this.Site.CaptureRequirement(
                                1049,
                                @"[In WeekdayOfMonth][The enumeration value]third[of the type WeekdayOfMonth means]Third");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1050
            this.Site.CaptureRequirement(
                                1050,
                                @"[In WeekdayOfMonth][The enumeration value]fourth[of the type WeekdayOfMonth means]Fourth");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1051
            this.Site.CaptureRequirement(
                                1051,
                                @"[In WeekdayOfMonth][The enumeration value]last[of the type WeekdayOfMonth means]Last");
        }

        #endregion

        #region MSOUTSPS_S02_TC29_OperationListItems_UIDUnique
        /// <summary>
        /// This test case is used to verify UID is unique among all other recurring events on this list.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC29_OperationListItems_UIDUnique()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();

            // Add 10 recurrence items with 10 unique UIDs' value.
            for (int recurrenceItemCounter = 1; recurrenceItemCounter <= 10; recurrenceItemCounter++)
            {
                string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
                Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(this.GetUniqueListItemTitle("DailyRecurrenceEvent"), eventDateOfRecurrence, "1", 10);
                this.Site.Assert.IsTrue(
                               settingsOfDailyRecurring.ContainsKey("UID"),
                               "The fields setting should contain UID field's setting.");

                this.Site.Assert.IsTrue(
                                     settingsOfDailyRecurring.ContainsKey("Title"),
                                     "The fields setting should contain Title field's setting.");

                settingsOfDailyRecurring["Title"] = eventTitle;
                settingsOfDailyRecurring["UID"] = Guid.NewGuid().ToString();
                addedItemsOfRecurrence.Add(settingsOfDailyRecurring);
                cmds.Add(MethodCmdEnum.New);
            }

            // Add 10 recurrence appointment items whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);

            // This method will verify there are should be 10 result element exist.
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 10);

            #endregion  Add one recurrence item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);

            // If passes upon verification, then capture R31202
            if (Common.IsRequirementEnabled(31202, this.Site))
            {
                // Verify UID value of each item is matched the value in fields setting.
                foreach (string addedRecurrenceItemid in addedRecurrenceItemIds)
                {
                    int zrowIndexOfCurrentId = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemid);
                    string actualUIDValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfCurrentId, "ows_UID");
                    this.Site.Assert.IsFalse(
                                        string.IsNullOrEmpty(actualUIDValue),
                                        "The list item[ID:{0}] should contain a UID value.",
                                        addedRecurrenceItemid);

                    Guid currentGUID;
                    if (!Guid.TryParse(actualUIDValue, out currentGUID))
                    {
                        this.Site.Assert.Fail("The UID field should be GUID format.Actual:[{0}]", actualUIDValue);
                    }

                    string formatedGuidValue = currentGUID.ToString();

                    // Verify whether there a fields setting of item which contain the actual UID value.
                    var fieldSettingsContainUIDValue = from fieldsSettingOfListItem in addedItemsOfRecurrence
                                                       where fieldsSettingOfListItem.Any(fieldSettingFounder => (fieldSettingFounder.Key.Equals("UID", StringComparison.OrdinalIgnoreCase) && fieldSettingFounder.Value.Equals(formatedGuidValue, StringComparison.OrdinalIgnoreCase)))
                                                       select fieldsSettingOfListItem;

                    this.Site.Assert.AreEqual<int>(
                                                1,
                                                fieldSettingsContainUIDValue.Count(),
                                                "There should be one match for UID value between the fields' setting and list item change data");
                }

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R31202
                this.Site.CaptureRequirement(
                                    31202,
                                    "[In Appendix B: Product Behavior] UID is unique for a sample of 10 among all other recurring events on this list.(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC30_OperationListItems_DurationValue
        /// <summary>
        /// This test case is used to verify Appointments have a duration value and an ending date and time
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC30_OperationListItems_DurationValue()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one single appointment item

            // Setting single setting
            DateTime eventDateOfSingleEvent = DateTime.Now.AddDays(1);
            DateTime endDateOfSingleEvent = eventDateOfSingleEvent.AddHours(1);

            string eventTitle = this.GetUniqueListItemTitle("SingleEvent");
            Dictionary<string, string> settingsOfSingleEvent = new Dictionary<string, string>();
            settingsOfSingleEvent.Add("Title", eventTitle);

            // Setting necessary fields' value
            settingsOfSingleEvent.Add("EventDate", this.GetGeneralFormatTimeString(eventDateOfSingleEvent));

            // The ending date and time of the appointment.
            settingsOfSingleEvent.Add("EndDate", this.GetGeneralFormatTimeString(endDateOfSingleEvent));

            // If the EventType indicates a single event 
            settingsOfSingleEvent.Add("EventType", "0");

            // Add single appointment.
            List<Dictionary<string, string>> addeditemsOfSingle = new List<Dictionary<string, string>>();
            addeditemsOfSingle.Add(settingsOfSingleEvent);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfSingle, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            #endregion  Add one single appointment item

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);

            if (Common.IsRequirementEnabled(146, this.Site))
            {
                int zrowIndexOfSingleAppointmentItem = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
                string actualDurationValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfSingleAppointmentItem, "ows_Duration");
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(actualDurationValue),
                                    "The Duration field should have value.");

                string actualEndDateValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfSingleAppointmentItem, "ows_EndDate");
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(actualEndDateValue),
                                    "The EndDate field should have value.");

                string actualEventDateValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfSingleAppointmentItem, "ows_EventDate");
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(actualEventDateValue),
                                    "The EventDate field should have value.");

                DateTime actualEventDate;
                if (!DateTime.TryParse(actualEventDateValue, out actualEventDate))
                {
                    this.Site.Assert.Fail("The EventDate field should have value.");
                }

                DateTime actualEndDate;
                if (!DateTime.TryParse(actualEndDateValue, out actualEndDate))
                {
                    this.Site.Assert.Fail("The EndDate field should have value.");
                }

                // If Duration, EventDate, EndDate have valid value, than capture R146
                TimeSpan expectedDurationTimeSpan = endDateOfSingleEvent - eventDateOfSingleEvent;

                double expectedDurationSecondsValue = expectedDurationTimeSpan.TotalSeconds;
                this.Site.Assert.AreEqual<string>(
                                expectedDurationSecondsValue.ToString(),
                                actualDurationValue,
                                "The actual Duration field value should equal to expected value.");

                // If duration value equal to the valid timeSpane value, than capture R13
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R13
                this.Site.CaptureRequirement(
                                        13,
                                        @"[In Appointments]Duration MUST be positive or zero.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R146
                this.Site.CaptureRequirement(
                                    146,
                                    @"[In Appendix B: Product Behavior] Implementation does have a duration value and an ending date and time.(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC31_OperationListItems_DeleteRecurrence
        /// <summary>
        /// This test case is used to verify When a recurrence is deleted, all exceptions to that recurrence also is deleted as optional behaviors.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC31_OperationListItems_DeleteRecurrence()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item.
            DateTime overwrittenRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwrittenRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            this.VerifyResponseOfUpdateListItem(updateResultOfException);
            this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            // Delete the recurrence item, but not specified the exception item id in the delete process.
            List<string> expectedDeletedListItems = new List<string>();
            expectedDeletedListItems.Add(recurrenceItemId);
            this.DeleteListItems(listId, expectedDeletedListItems, OnErrorEnum.Continue);

            // Get the list items change data.
            XmlNode[] zrowItems = this.TryGetListItemsChangesFromSUT(listId, null);

            if (Common.IsRequirementEnabled(36002, this.Site))
            {
                // If the rs:date element does not contain any zrow items, that means the exception list item id does not present in the response of GetListItemChangesSinceToken operation, than capture R36002.
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R36002
                this.Site.CaptureRequirementIfAreEqual<int>(
                                              0,
                                              zrowItems.Length,
                                              36002,
                                              @"[In Appendix B: Product Behavior][In Appendix B: Product Behavior]In Appendix B: Product Behavior] Implementation does also be deleted.(Microsoft® Office Outlook® 2003, Microsoft® Office Outlook® 2007, Microsoft® Outlook® 2010, Microsoft® Outlook® 2013, Microsoft® SharePoint® Foundation 2010, Microsoft® SharePoint® Foundation 2013 products follow this behavior)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC32_OperationListItemsForAppointment_TimeZone
        /// <summary>
        /// This test case is used to verify if fRecurrence is TRUE, TimeZone contains an integer index into a list of time zones.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC32_OperationListItemsForAppointment_TimeZone()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();

            this.Site.Assert.AreEqual<string>(
                                    "1",
                                    settingsOfDailyRecurring["fRecurrence"].ToLower(),
                                    "The fRecurrence field should be '1' in this test case.");

            // Setting the TimeZone field.
            int timeZoneIdOfPacificTime = Common.GetConfigurationPropertyValue<int>("TimeZoneIDOfPacificTime", this.Site);

            if (settingsOfDailyRecurring.ContainsKey("TimeZone"))
            {
                settingsOfDailyRecurring["TimeZone"] = timeZoneIdOfPacificTime.ToString();
            }
            else
            {
                settingsOfDailyRecurring.Add("TimeZone", timeZoneIdOfPacificTime.ToString());
            }

            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Set "<ViewFields />" in order to show all fields' value of a list.
                CamlViewFields viewFields = new CamlViewFields();
                viewFields.ViewFields = new CamlViewFieldsViewFields();

                // Call GetListItemChangesSinceToken operation to get list items change.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                         listId,
                                         null,
                                         null,
                                         viewFields,
                                         null,
                                         null,
                                         null,
                                         null);

                // Get the list items change data.
                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                if (Common.IsRequirementEnabled(442, this.Site))
                {
                    int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
                    string timeZoneValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_TimeZone");
                    this.Site.Assert.IsFalse(
                                           string.IsNullOrEmpty(timeZoneValue),
                                            "The TimeZone field value should have value.");

                    this.Site.Assert.AreEqual<string>(
                                            timeZoneIdOfPacificTime.ToString(),
                                            timeZoneValue,
                                            "The TimeZone value in response of GetListItemChangesSinceToken operation should equal to set in below step.");

                    // Verify MS-OUTSPS requirement: MS-OUTSPS_R442
                    this.Site.CaptureRequirement(
                                              442,
                                              @"[In Appendix B: Product Behavior] Implementation does contain an integer index into a list of time zones.(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                }
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC33_OperationListItems_TimeZoneSetByProtocolSUT
        /// <summary>
        /// This test case is used to verify TimeZone doesn’t be empty.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC33_OperationListItems_TimeZoneSetByProtocolSUT()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();

            this.Site.Assert.AreEqual<string>(
                                    "1",
                                    settingsOfDailyRecurring["fRecurrence"].ToLower(),
                                    "The fRecurrence field should be '1' in this test case.");

            this.Site.Assert.IsFalse(
                                settingsOfDailyRecurring.ContainsKey("TimeZone"),
                                "The TimeZone field should absent from the request in this test case.");

            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Set "<ViewFields />" in order to show all fields' value of a list.
                CamlViewFields viewFields = new CamlViewFields();
                viewFields.ViewFields = new CamlViewFieldsViewFields();

                // Call GetListItemChangesSinceToken operation to get list items change.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                         listId,
                                         null,
                                         null,
                                         viewFields,
                                         null,
                                         null,
                                         null,
                                         null);

                // Get the list items change data.
                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);
                int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
                string timeZoneValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_TimeZone");

                if (Common.IsRequirementEnabled(1270, this.Site))
                {
                    // If the TimeZone field value is set by protocol SUT, and it is valid integer value, than capture R1270
                    if (!string.IsNullOrEmpty(timeZoneValue))
                    {
                        int timeZoneIntValue;
                        if (!int.TryParse(timeZoneValue, out timeZoneIntValue))
                        {
                            this.Site.Assert.Fail("The TimeZone field value should be valid integer value.");
                        }

                        // Verify MS-OUTSPS requirement: MS-OUTSPS_R1270
                        this.Site.CaptureRequirement(
                                                  1270,
                                                  @"[In Appendix B: Product Behavior] [If does not leave it empty]Implementation does set a number in this value[TimeZone].(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                    }
                }
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC34_TriggerExceptionDeletion_UpdateEndDate
        /// <summary>
        /// This test case is used to verify protocol servers will trigger exception deletion when EndDate is updated.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC34_TriggerExceptionDeletion_UpdateEndDate()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item.
            DateTime overwrittenRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwrittenRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            #region Update the EndDate field to trigger exception deletion.

            // Set the EventDate different than the event date set before.
            string orginalEndTimeStringValue = settingsOfDailyRecurring["EndDate"];
            DateTime orginalEndTime;
            if (!DateTime.TryParse(orginalEndTimeStringValue, out orginalEndTime))
            {
                this.Site.Assert.Fail(
                            "The EndDate field should be valid DateTime format.Actual:[{0}]",
                            settingsOfDailyRecurring["EndDate"]);
            }

            DateTime updatedEndDateOfRecurrence = orginalEndTime.AddHours(1);
            settingsOfDailyRecurring["EndDate"] = this.GetUTCFormatTimeString(updatedEndDateOfRecurrence);

            // Append "deleteExceptions" element on RecurrenceData field to tell Protocol SUT should trigger exception deletion.
            this.AppenddeleteExceptionsElement(settingsOfDailyRecurring);

            List<MethodCmdEnum> updateRecurrencecmds = new List<MethodCmdEnum>(1);
            updateRecurrencecmds.Add(MethodCmdEnum.Update);

            // Set target update list item
            settingsOfDailyRecurring.Add("ID", recurrenceItemId);

            // Update the recurrence item with updated RecurrenceData field value.
            List<Dictionary<string, string>> updatedRecurrenceSettings = new List<Dictionary<string, string>>();
            updatedRecurrenceSettings.Add(settingsOfDailyRecurring);
            UpdateListItemsUpdates updatesOfDeletion = this.CreateUpdateListItems(updateRecurrencecmds, updatedRecurrenceSettings, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfTriggerExceptionDeletion = OutspsAdapter.UpdateListItems(listId, updatesOfDeletion);
            this.VerifyResponseOfUpdateListItem(updateResultOfTriggerExceptionDeletion);

            #endregion Update the EventDate field to trigger exception deletion.

            // If above UpdateListItems operations perform successfully, then capture R1069
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1069
            this.Site.CaptureRequirement(
                                1069,
                                @"[In Message Processing Events and Sequencing Rules][The operation]UpdateListItems Creates or modifies items on a list.");

            XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetListItemsChangesFromSUT(listId);

            // If the exception item data does not present in zrow items array in response of GetListItemChangesSinceToken, then capture R445021
            int zrowIndexOfExceptionItem = this.TryGetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, exceptionItemIds[0]);

            if (Common.IsRequirementEnabled(445021, this.Site))
            {
                this.Site.Assert.AreEqual<int>(
                                      -1,
                                      zrowIndexOfExceptionItem,
                                      "The exception item data should absent in zrow items array in response of GetListItemChangesSinceToken");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R445021
                this.Site.CaptureRequirement(
                                          445021,
                                          @"[In Appendix B: Product Behavior] Implementation does trigger exception deletion when EndDate is updated.(<27>Windows SharePoint Services3.0 and above does delete exception items when these properties are updated.)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC35_TriggerExceptionDeletion_UpdateEventDate
        /// <summary>
        /// This test case is used to verify protocol servers will trigger exception deletion when EventDate is updated.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC35_TriggerExceptionDeletion_UpdateEventDate()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item.
            DateTime overwrittenRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwrittenRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            #region Update the EventDate field to trigger exception deletion.

            // Set the EventDate different than the event date set before.
            DateTime updatedEventDateOfRecurrence = eventDateOfRecurrence.AddHours(0.5);
            settingsOfDailyRecurring["EventDate"] = this.GetUTCFormatTimeString(updatedEventDateOfRecurrence);

            // Append "deleteExceptions" element on RecurrenceData field to tell Protocol SUT should trigger exception deletion.
            this.AppenddeleteExceptionsElement(settingsOfDailyRecurring);

            List<MethodCmdEnum> updateRecurrencecmds = new List<MethodCmdEnum>(1);
            updateRecurrencecmds.Add(MethodCmdEnum.Update);

            // Set target update list item
            settingsOfDailyRecurring.Add("ID", recurrenceItemId);

            // Update the recurrence item with updated RecurrenceData field value.
            List<Dictionary<string, string>> updatedRecurrenceSettings = new List<Dictionary<string, string>>();
            updatedRecurrenceSettings.Add(settingsOfDailyRecurring);
            UpdateListItemsUpdates updatesOfDeletion = this.CreateUpdateListItems(updateRecurrencecmds, updatedRecurrenceSettings, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfTriggerExceptionDeletion = OutspsAdapter.UpdateListItems(listId, updatesOfDeletion);
            this.VerifyResponseOfUpdateListItem(updateResultOfTriggerExceptionDeletion);

            #endregion Update the EventDate field to trigger exception deletion.

            XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetListItemsChangesFromSUT(listId);

            // If the exception item data does not present in zrow items array in response of GetListItemChangesSinceToken, then capture R445022
            int zrowIndexOfExceptionItem = this.TryGetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, exceptionItemIds[0]);

            if (Common.IsRequirementEnabled(445022, this.Site))
            {
                this.Site.Assert.AreEqual<int>(
                                    -1,
                                    zrowIndexOfExceptionItem,
                                    "The exception item data should absent in zrow items array in response of GetListItemChangesSinceToken");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R445022
                this.Site.CaptureRequirement(
                                          445022,
                                          @"[In Appendix B: Product Behavior] Implementation does trigger exception deletion when one of these properties EventDate is updated.(<27>Windows SharePoint Services3.0 and above does delete exception items when these properties are updated.)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC36_TriggerExceptionDeletion_UpdateRecurrenceData
        /// <summary>
        /// This test case is used to verify protocol servers will trigger exception deletion when RecurrenceData is updated.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC36_TriggerExceptionDeletion_UpdateRecurrenceData()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item.
            DateTime overwrittenRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwrittenRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            #region Update the RecurrenData field to trigger exception deletion.

            // Update the RecurrenceData field and append the deleteExceptions element.
            RecurrenceXML recurrenceXMLData = new RecurrenceXML();

            // recurrenceXMLData.deleteExceptions = null;
            recurrenceXMLData.recurrence = new RecurrenceDefinition();
            recurrenceXMLData.recurrence.rule = new RecurrenceRule();
            recurrenceXMLData.recurrence.rule.firstDayOfWeek = DayOfWeekOrMonth.mo;

            // Set the windowEnd value, the recurrence will be end on one month later, and the repeat frequency is weekly.
            DateTime winEndDate = eventDateOfRecurrence.Date.AddDays(30);
            recurrenceXMLData.recurrence.rule.Item = winEndDate;

            // Repeat Pattern, set it as weekly repeat.
            recurrenceXMLData.recurrence.rule.repeat = new RepeatPattern();

            // Setting the WeeklyFrequencyValue
            RepeatPatternWeekly repeatPatternWeeklyData = new RepeatPatternWeekly();
            repeatPatternWeeklyData.weekFrequency = "1";
            repeatPatternWeeklyData.mo = TrueFalseDOW.TRUE;
            repeatPatternWeeklyData.moSpecified = true;
            recurrenceXMLData.recurrence.rule.repeat.Item = repeatPatternWeeklyData;

            // Update the RecurrenData field
            string recurrenceXMLString = this.GetRecurrenceXMLString(recurrenceXMLData);
            settingsOfDailyRecurring["RecurrenceData"] = recurrenceXMLString;

            List<MethodCmdEnum> updateRecurrencecmds = new List<MethodCmdEnum>(1);
            updateRecurrencecmds.Add(MethodCmdEnum.Update);

            // Set target update list item
            settingsOfDailyRecurring.Add("ID", recurrenceItemId);

            // Update the recurrence item with updated RecurrenceData field value.
            List<Dictionary<string, string>> updatedRecurrenceSettings = new List<Dictionary<string, string>>();
            updatedRecurrenceSettings.Add(settingsOfDailyRecurring);
            UpdateListItemsUpdates updatesOfDeletion = this.CreateUpdateListItems(updateRecurrencecmds, updatedRecurrenceSettings, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfTriggerExceptionDeletion = OutspsAdapter.UpdateListItems(listId, updatesOfDeletion);
            this.VerifyResponseOfUpdateListItem(updateResultOfTriggerExceptionDeletion);

            #endregion Update the RecurrenData field to trigger exception deletion.

            XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetListItemsChangesFromSUT(listId);

            // If the exception item data does not present in zrow items array in response of GetListItemChangesSinceToken, then capture R445023, R900
            int zrowIndexOfExceptionItem = this.TryGetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, exceptionItemIds[0]);

            if (Common.IsRequirementEnabled(445023, this.Site))
            {
                this.Site.Assert.AreEqual<int>(
                                        -1,
                                        zrowIndexOfExceptionItem,
                                        "The exception item data should absent in zrow items array in response of GetListItemChangesSinceToken");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R445023
                this.Site.CaptureRequirement(
                                          445023,
                                          @"[In Appendix B: Product Behavior] Implementation does trigger exception deletion when RecurrenceData is updated.(<27>Windows SharePoint Services3.0 and above does delete exception items when these properties are updated.)");
            }

            this.Site.Assert.AreEqual<int>(
                                       -1,
                                       zrowIndexOfExceptionItem,
                                       "The exception item data should absent in zrow items array in response of GetListItemChangesSinceToken");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R900
            this.Site.CaptureRequirement(
                                      900,
                                      @"[In RecurrenceXML complex type]deleteExceptions: This element MUST be present if and only if RecurrenceXML is written by a protocol client in UpdateListItems (section 3.1.4.10) and the protocol client requests that the protocol server delete all exception items for this recurrence. See section 3.2.1.1 for details about exception items and recurrences.");
        }

        #endregion

        #region MSOUTSPS_S02_TC37_TriggerExceptionDeletion_UpdateXMLTZone
        /// <summary>
        /// This test case is used to verify protocol servers will trigger exception deletion when XMLTZone is updated.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC37_TriggerExceptionDeletion_UpdateXMLTZone()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            #region  Add one recurrence item

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
            string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
            Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
            List<Dictionary<string, string>> addedItemsOfRecurrence = new List<Dictionary<string, string>>();
            addedItemsOfRecurrence.Add(settingsOfDailyRecurring);

            // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addedItemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

            // Validate the RecurrenceData Field
            string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
            this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

            #endregion  Add one recurrence item

            #region Add exception item
            // add an exception appointment item whose event day is different from recurrence item.
            DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

            // The second recurrence instance will be replaced by exception item.
            DateTime overwrittenRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

            // Get the recurrence item id.
            string recurrenceItemId = addedRecurrenceItemIds[0];

            // Set the exception item setting.
            string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
            Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                exceptionEventDate,
                                                                                overwrittenRecerrenceEventDate,
                                                                                exceptionItemTitle,
                                                                                recurrenceItemId,
                                                                                settingsOfDailyRecurring);

            List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
            addeditemsOfException.Add(settingsOfException);
            UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

            // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
            List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

            #endregion Add exception item

            #region Update the XMLTZone field to trigger exception deletion.

            // Update the XMlTZone to (-480) minutes offset from UTC Time (UTC = local time + bias), current TimeZone is UTC+ 8
            settingsOfDailyRecurring["XMLTZone"] = @"<timeZoneRule><standardBias>-480</standardBias></timeZoneRule>";

            // Setting the target list item id, it is the recurrence item.
            settingsOfDailyRecurring["ID"] = addedRecurrenceItemIds[0];

            // Append "deleteExceptions" element on RecurrenceData field to tell Protocol SUT should trigger exception deletion.
            this.AppenddeleteExceptionsElement(settingsOfDailyRecurring);

            // Update the recurrence item with updated event date value.
            List<Dictionary<string, string>> updatedRecurrenceSettings = new List<Dictionary<string, string>>();
            updatedRecurrenceSettings.Add(settingsOfDailyRecurring);

            List<MethodCmdEnum> updateRecurrencecmds = new List<MethodCmdEnum>();
            updateRecurrencecmds.Add(MethodCmdEnum.Update);
            UpdateListItemsUpdates updatesOfDeletion = this.CreateUpdateListItems(updateRecurrencecmds, updatedRecurrenceSettings, OnErrorEnum.Continue);

            // Call UpdateListItems to update the XMLTZone field for the recurrence, so that the protocol SUT will trigger the exception items' deletion.
            UpdateListItemsResponseUpdateListItemsResult updateResultOfTriggerExcetionDeletion = OutspsAdapter.UpdateListItems(listId, updatesOfDeletion);
            this.VerifyResponseOfUpdateListItem(updateResultOfTriggerExcetionDeletion);

            #endregion Update the XMLTZone field to trigger exception deletion.

            XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetListItemsChangesFromSUT(listId);

            // If the exception item data does not present in zrow items array in response of GetListItemChangesSinceToken, then capture R445024
            int zrowIndexOfExceptionItem = this.TryGetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, exceptionItemIds[0]);

            if (Common.IsRequirementEnabled(445024, this.Site))
            {
                this.Site.Assert.AreEqual<int>(
                                   -1,
                                   zrowIndexOfExceptionItem,
                                   "The exception item data should absent in zrow items array in response of GetListItemChangesSinceToken");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R445024
                this.Site.CaptureRequirement(
                                          445024,
                                          @"[In Appendix B: Product Behavior] Implementation does trigger exception deletion when XMLTZone is updated.(<27>Windows SharePoint Services3.0 and above does delete exception items when these properties are updated.)");
            }
        }
        #endregion

        #region MSOUTSPS_S02_TC38_GetListItemChangesSinceToken_Support
        /// <summary>
        /// This test case is used to verify the value of the server version is "12.0.0.4326" or greater indicates the server supports GetListItemChangesSinceToken.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC38_GetListItemChangesSinceToken_Support()
        {
            // Add a list and add one item.
            string listId = this.AddListToSUT(TemplateType.Generic_List);
            List<string> listItemsIds = this.AddItemsToList(listId, 1);

            // Setting "<viewFields />" to view all fields of the list.
            CamlViewFields viewfieds = new CamlViewFields();
            viewfieds.ViewFields = new CamlViewFieldsViewFields();

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Call GetListItemChangesSinceToken operation to get items' change.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId,
                                            null,
                                            null,
                                            viewfieds,
                                            null,
                                            null,
                                            null,
                                            null);

                // Get the list items change data.
                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                // If the zrow index can be get and no Assert exception is thrown by GetZrowItemIndexByListItemId method, that means the added item id is present in response of GetListItemChangesSinceToken operation, then capture R12552, R106802, R12160
                this.GetZrowItemIndexByListItemId(zrowItems, listItemsIds[0]);

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R106802
                this.Site.CaptureRequirement(
                                    106802,
                                    @"[In Appendix B: Product Behavior] Implementation does support GetListItemChangesSinceToke.(Microsoft® Office Outlook® 2003 and Windows SharePoint Services 3.0 and above follow this behavior)");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R12160
                this.Site.CaptureRequirement(
                                    12160,
                                    @"[In Messages]GetListItemChangesSinceTokenResponse specified the response to a request to download changes that have happened since the protocol client's last download on any protocol server that supports it.");

                if (Common.IsRequirementEnabled(12552, this.Site))
                {
                    #region Get protocol SUT version and verify it

                    GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);
                    if (null == getListResult || null == getListResult.List)
                    {
                        this.Site.Assert.Fail("The response of GetList operation should contain valid List element.");
                    }

                    string sutVersion = getListResult.List.ServerSettings.ServerVersion;
                    this.Site.Assert.IsFalse(
                                        string.IsNullOrEmpty(sutVersion),
                                        "The ServerVersion element should have value.");

                    bool isCurrentVersionEqualOrLargerThanSpecfied = VerifyEqualOrLargerThanSpecifiedVersionString(sutVersion, "12.0.0.4326");

                    #endregion Get protocol SUT version and verify it

                    // If the server version value is equal or larger than "12.0.0.4326" and the protocol SUT support the GetListItemChangesSinceToken operation, then capture R12552.
                    this.Site.Assert.IsTrue(
                                        isCurrentVersionEqualOrLargerThanSpecfied,
                                        "The server version should be equal or larger than [12.0.0.4326], that indicate the protocol SUT support the GetListItemChangesSinceToken operation.");

                    // Verify MS-OUTSPS requirement: MS-OUTSPS_R12552
                    this.Site.CaptureRequirement(
                                        12552,
                                        @"[In Appendix B: Product Behavior][<2> Section 3.1.4: ]A value of ""12.0.0.4326"" or greater indicates the server supports GetListItemChangesSinceToken.");
                }
            }

            // Call GetListItemChanges operation to get items' change.
            GetListItemChangesResponseGetListItemChangesResult responseOfGetListItemChanges = null;
            responseOfGetListItemChanges = OutspsAdapter.GetListItemChanges(
                                                                        listId,
                                                                        viewfieds,
                                                                        null,
                                                                        null);

            this.Site.Assert.IsNotNull(
                                    responseOfGetListItemChanges,
                                    "The response of GetListItemChangesSinceToken operation should have value.");
            this.Site.Assert.IsNotNull(
                         responseOfGetListItemChanges.listitems.data,
                         "The response of GetListItemChangesSinceToken operation should contain [zrow] data structure under [listitems] element.");

            XmlNode[] zrowItemsOfGetListItemChanges = this.GetZrowItems(responseOfGetListItemChanges.listitems.data[0].Any);

            // If the zrow index can be get and no Assert exception is thrown by GetZrowItemIndexByListItemId method, that means the added item id is present in response of GetListItemChangesSinceToken operation.
            this.GetZrowItemIndexByListItemId(zrowItemsOfGetListItemChanges, listItemsIds[0]);
        }

        #endregion

        #region MSOUTSPS_S02_TC39_GetListItemChangesSinceToken_QueryIsEmpty
        /// <summary>
        /// This test case is used to verify in Windows SharePoint Services 3.0 and SharePoint Foundation 2010, 
        /// if there is no "query" element in the request, the server will sort items by the ID field, in ascending order.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC39_GetListItemChangesSinceToken_QueryIsEmpty()
        {
            string listId = this.AddListToSUT(TemplateType.Generic_List);

            // Add 10 list items
            int addedItemNumber = 10;
            this.AddItemsToList(listId, addedItemNumber);

            // Get list items' changes from protocol SUT
            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);

            if (Common.IsRequirementEnabled(11420, this.Site))
            {
                #region verify whether sort by ascending

                List<string> listItemsIds = new List<string>();
                int previousListItemIntValue = -1;
                for (int arrayIndex = 0; arrayIndex < zrowItems.Length; arrayIndex++)
                {
                    var listItemIdItems = from XmlAttribute attributeItem in zrowItems[arrayIndex].Attributes
                                          where attributeItem.Name.Equals("ows_ID", StringComparison.OrdinalIgnoreCase)
                                          select attributeItem.Value;

                    if (listItemIdItems.Count() != 1)
                    {
                        this.Site.Assert.Fail("Could not get the valid ID field value.");
                    }
                    else
                    {
                        string listItemIdStringValue = listItemIdItems.ElementAt<string>(0);
                        int listItemIdIntValue;
                        if (!int.TryParse(listItemIdStringValue, out listItemIdIntValue))
                        {
                            this.Site.Assert.Fail("Each list item id should be integer format");
                        }

                        this.Site.Assert.IsTrue(
                                            listItemIdIntValue > 0,
                                            "The list item id should large than zero.");

                        if (listItemIdIntValue == previousListItemIntValue)
                        {
                            continue;
                        }
                        else if (listItemIdIntValue < previousListItemIntValue)
                        {
                            this.Site.Assert.Fail(
                                        "The list item id should sort by ascending.\r\n Position[{0}]: [{1}]\r\n Position[{2}]: [{3}]",
                                        0 == arrayIndex ? 0 : arrayIndex - 1,
                                        previousListItemIntValue,
                                        arrayIndex,
                                        listItemIdIntValue);
                        }
                        else
                        {
                            listItemsIds.Add(listItemIdIntValue.ToString());
                            previousListItemIntValue = listItemIdIntValue;
                        }
                    }
                }

                #endregion verify whether sort by ascending

                // If the zrow items are sort by ascending, then capture R11420.
                this.Site.Assert.AreEqual<int>(
                                addedItemNumber,
                                listItemsIds.Count(),
                                "The response of GetListItemChangesSinceToken should contain all added list items. Total:[{0}]",
                                addedItemNumber);

                // Since the zrow items are sort by ascending, capture R11420.
                this.Site.CaptureRequirement(
                                11420,
                                @"[In Appendix B: Product Behavior]Implementation does sort items by the ID field, in ascending order. (<8> Section 3.1.4.7: In Windows SharePoint Services 3.0 and SharePoint Foundation 2010, if there is no ""query"" element in the request, the server will sort items by the ID field, in ascending order.)");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC40_GetListItemChangesSinceToken_HaveInstances
        /// <summary>
        /// This test case is used to verify the instance of a recurrence can have zero or one total exceptions and deleted instances
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC40_GetListItemChangesSinceToken_HaveInstances()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            // If upon verification pass, then capture R8120
            if (Common.IsRequirementEnabled(8120, this.Site))
            {
                #region Add one recurrence item

                List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
                cmds.Add(MethodCmdEnum.New);

                // Setting recurring setting
                DateTime eventDateOfRecurrence = DateTime.Today.Date.AddDays(1);
                string eventTitle = this.GetUniqueListItemTitle("DailyRecurrenceEvent");
                Dictionary<string, string> settingsOfDailyRecurring = this.GetDailyRecurrenceSettingWithwindowEnd(eventTitle, eventDateOfRecurrence, "1", 10);
                List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
                addeditemsOfRecurrence.Add(settingsOfDailyRecurring);

                // add a  recurrence appointment item whose duration is 10 days, and each instance of this recurrence item will be repeat on "Daily". There should be 10 instances of this item.
                UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

                // Get list item id from the response of UpdateListItems operation.
                List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
                XmlNode[] zrowItemsOfRecurrenceAppointment = this.GetZrowItems(updateResult.Results[0].Any);

                // Validate the RecurrenceData Field
                string updatedRecurrenceDataValue = Common.GetZrowAttributeValue(zrowItemsOfRecurrenceAppointment, 0, "ows_RecurrenceData");
                this.VerifyComplexTypesSchema(updatedRecurrenceDataValue, typeof(RecurrenceXML));

                #endregion Add one recurrence item

                #region Add exception item
                // add an exception appointment item whose event day is different from recurrence item.
                DateTime exceptionEventDate = eventDateOfRecurrence.AddDays(1).AddHours(2);

                // The second recurrence instance will be replaced by exception item.
                DateTime overwritedRecerrenceEventDate = eventDateOfRecurrence.AddDays(1);

                // Get the recurrence item id.
                string recurrenceItemId = addedRecurrenceItemIds[0];

                // Set the exception item setting.
                string exceptionItemTitle = this.GetUniqueListItemTitle("ExceptionItem");
                Dictionary<string, string> settingsOfException = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                    exceptionEventDate,
                                                                                    overwritedRecerrenceEventDate,
                                                                                    exceptionItemTitle,
                                                                                    recurrenceItemId,
                                                                                    settingsOfDailyRecurring);

                List<Dictionary<string, string>> addeditemsOfException = new List<Dictionary<string, string>>();
                addeditemsOfException.Add(settingsOfException);
                UpdateListItemsUpdates updatesOfException = this.CreateUpdateListItems(cmds, addeditemsOfException, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResultOfException = OutspsAdapter.UpdateListItems(listId, updatesOfException);

                // Get the exception item Id, and only add 1 exception item in this UpdateListItems call.
                List<string> exceptionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfException, 1);

                #endregion Add exception item

                #region Add deletion item

                // Add a deletion appointment item whose event day is replace the third instance of the recurrence item.
                DateTime deletionEventDate = eventDateOfRecurrence.AddDays(2).AddHours(2);

                // The third recurrence instance will be replaced by deletion item.
                DateTime overwritedRecerrenceEventDateByDeletionItem = eventDateOfRecurrence.AddDays(2);

                // Set the deletion item setting.
                string deletionItemTitle = this.GetUniqueListItemTitle("DeletionItem");
                Dictionary<string, string> settingsOfDeletion = this.GetExceptionsItemSettingForRecurrenceEvent(
                                                                                    deletionEventDate,
                                                                                    overwritedRecerrenceEventDateByDeletionItem,
                                                                                    deletionItemTitle,
                                                                                    recurrenceItemId,
                                                                                    settingsOfDailyRecurring);
                this.Site.Assert.IsTrue(
                            settingsOfDeletion.ContainsKey("EventType"),
                            "The settings Of deletion should contain the 'EventType' value.");

                // The only difference between the exception and deletion item is the "EventType" field value. For deletion item, set it as "3".
                settingsOfDeletion["EventType"] = "3";

                List<Dictionary<string, string>> addeditemsOfDeletion = new List<Dictionary<string, string>>();
                addeditemsOfDeletion.Add(settingsOfDeletion);
                UpdateListItemsUpdates updatesOfDeletion = this.CreateUpdateListItems(cmds, addeditemsOfDeletion, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResultOfDeletion = OutspsAdapter.UpdateListItems(listId, updatesOfDeletion);

                // Get the deletion item Id, and only add 1 deletion item in this UpdateListItems call.
                List<string> deletionItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResultOfDeletion, 1);

                #endregion Add deletion recurrence item

                #region  Add one single appointment item

                // Setting single setting
                DateTime eventDateOfSingleEvent = DateTime.Now.AddDays(1);
                DateTime endDateOfSingleEvent = eventDateOfSingleEvent.AddHours(1);

                string eventTitleOfSingleEvent = this.GetUniqueListItemTitle("SingleEvent");
                Dictionary<string, string> settingsOfSingleEvent = new Dictionary<string, string>();

                // Setting necessary fields' value
                settingsOfSingleEvent.Add("EventDate", this.GetGeneralFormatTimeString(eventDateOfSingleEvent));

                // The ending date and time of the appointment.
                settingsOfSingleEvent.Add("EndDate", this.GetGeneralFormatTimeString(endDateOfSingleEvent));

                // If the EventType indicates a single event 
                settingsOfSingleEvent.Add("EventType", "0");

                settingsOfSingleEvent.Add("Title", eventTitleOfSingleEvent);

                // Add single appointment.
                List<Dictionary<string, string>> addeditemsOfSingleEvent = new List<Dictionary<string, string>>();
                addeditemsOfSingleEvent.Add(settingsOfSingleEvent);
                List<MethodCmdEnum> cmdsOfSingleEvent = new List<MethodCmdEnum>(1);
                cmdsOfSingleEvent.Add(MethodCmdEnum.New);

                UpdateListItemsUpdates updatesOfRecurrenceOfSingleItem = this.CreateUpdateListItems(cmdsOfSingleEvent, addeditemsOfSingleEvent, OnErrorEnum.Continue);
                UpdateListItemsResponseUpdateListItemsResult updateResultOfSingleItem = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrenceOfSingleItem);
                this.VerifyResponseOfUpdateListItem(updateResultOfSingleItem);

                // Get list item id from the response of UpdateListItems operation.
                this.VerifyResponseOfUpdateListItem(updateResult);
                List<string> addedRecurrenceItemIdsOfSingleItem = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

                #endregion  Add one single appointment item

                XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetListItemsChangesFromSUT(listId);

                this.Site.Assert.AreEqual<int>(
                                4,
                                zrowItemsOfGetListItemChangesSinceToken.Count(),
                                "The zrow items should equal to the number of added list items[3]:Recurrence item[{0}]\r\n, exception item[{1}]\r\n, deletion item[{2}]\r\n.",
                                eventTitle,
                                exceptionItemTitle,
                                deletionItemTitle);

                List<string> addedListitemIds = new List<string>();
                addedListitemIds.Add(recurrenceItemId);
                addedListitemIds.Add(exceptionItemIds[0]);
                addedListitemIds.Add(deletionItemIds[0]);
                addedListitemIds.Add(addedRecurrenceItemIdsOfSingleItem[0]);

                // Verify all added list item ids are in the zrow items' array.
                foreach (string addedListItemIdItem in addedListitemIds)
                {
                    bool isListItemIdIncludedInZrowItem = this.VerifyContainExpectedListItemById(addedListItemIdItem, zrowItemsOfGetListItemChangesSinceToken);
                    this.Site.Assert.IsTrue(
                                        isListItemIdIncludedInZrowItem,
                                        "The list item id[{0}] should be included in zrow items' array.",
                                        addedListItemIdItem);
                }

                // If the eventType field value of the deletion item equal to 3, then capture R1011
                int zrowIndexOfDeletionItem = this.GetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, deletionItemIds[0]);
                string eventTypeValue = Common.GetZrowAttributeValue(zrowItemsOfGetListItemChangesSinceToken, zrowIndexOfDeletionItem, "ows_EventType");

                // Verify the eventType schema definition.If there are no any schema validation, capture R274, R16
                for (int zrowIndexTemp = 0; zrowIndexTemp < zrowItemsOfGetListItemChangesSinceToken.Length; zrowIndexTemp++)
                {
                    string eventTypeValueTemp = Common.GetZrowAttributeValue(zrowItemsOfGetListItemChangesSinceToken, zrowIndexTemp, "ows_EventType");
                    this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(eventTypeValueTemp),
                                    "Each zrow item should have EventType field value.");
                }

                this.VerifySimpleTypeSchema(zrowItemsOfGetListItemChangesSinceToken);

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R274
                this.Site.CaptureRequirement(
                                        274,
                                        @"[In Appointment-Specific Schema]EventType: An EventType (section 2.2.5.5) integer.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R16
                this.Site.CaptureRequirement(
                                        16,
                                        @"[In Appointments]Appointments MUST be one of four types: Single,Recurring,an exception to a recurrence or a deleted instance of a recurrence.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1011
                this.Site.CaptureRequirementIfAreEqual<string>(
                                            "3",
                                            eventTypeValue,
                                            1011,
                                            @"[In EventType][The enumeration value]3[of the type EventType means]Deleted instance of a recurrence.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R8120
                this.Site.CaptureRequirement(
                                        8120,
                                        @"[In Appendix B: Product Behavior] Implementation does have zero or one total exceptions and deleted instances. (<12> Section 3.2.1.1.2:  Windows SharePoint Services 2.0, Windows SharePoint Services 3.0, and SharePoint Foundation 2010 allow this).");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC41_AddDiscussionBoardItem
        /// <summary>
        /// This test case is used to verify AddDiscussionBoardItem operation to add a new discussion item to a list.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC41_AddDiscussionBoardItem()
        {
            // Add a discussion board list.
            string listId = this.AddListToSUT(TemplateType.Discussion_Board);
            byte[] messageData = this.GetMessageDataForAddDiscussionBoardItem();

            // Call AddDiscussionBoardItem operation to add a discussionBoard item
            AddDiscussionBoardItemResponseAddDiscussionBoardItemResult responseOfAddDiscussionBoardItem = null;
            responseOfAddDiscussionBoardItem = OutspsAdapter.AddDiscussionBoardItem(listId, messageData);
            this.Site.Assert.IsNotNull(
                         responseOfAddDiscussionBoardItem.listitems.data,
                         "The response of AddDiscussionBoardItem operation should contain [zrow] data structure under [listitems] element.");

            this.Site.Assert.IsNotNull(
                          responseOfAddDiscussionBoardItem.listitems.data.Any,
                          "The response of AddDiscussionBoardItem operation should contain at least one zrow item.");

            XmlNode[] zrowItems = this.GetZrowItems(responseOfAddDiscussionBoardItem.listitems.data.Any);
            this.Site.Assert.AreEqual<int>(
                             1,
                             zrowItems.Count(),
                             "The response of AddDiscussionBoardItem operation should contain only one zrow item.");

            string listItemIdOfDiscussionBoardItem = Common.GetZrowAttributeValue(zrowItems, 0, "ows_ID");
            this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listItemIdOfDiscussionBoardItem),
                                "The response of AddDiscussionBoardItem operation should contain the Id field value.");

            // If the response of AddDiscussionBoardItem operation succeed, then capture R10860, R1088
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R10860
            this.Site.CaptureRequirement(
                            10860,
                            @"[In Messages]AddDiscussionBoardItemResponse specified the response to a request to add new discussion items to a list.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1088
            this.Site.CaptureRequirement(
                             1088,
                             @"[In AddDiscussionBoardItemResponse]The item identifier is found in the attribute AddDiscussionBoardItemResponse.AddDiscussionBoardItemResult.listitems.data.row.ows_ID.");

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                        listId,
                                                        null,
                                                        null,
                                                        viewfieds,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItemsOfGetListItemChangesSinceToken = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                // Current discussion board list only contain 1 item, so the list item id should be 1. The zrow items should contain the added list item data.
                // If there are no any match zrow items found in zrow items array, this method will throw Assert exception.
                this.GetZrowItemIndexByListItemId(zrowItemsOfGetListItemChangesSinceToken, "1");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1063
                this.Site.CaptureRequirement(
                                 1063,
                                 @"[In Message Processing Events and Sequencing Rules][The operation]AddDiscussionBoardItem Adds a new discussion item to a list.");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC42_GetListItemChangesSinceToken_OptimizeLookups
        /// <summary>
        /// This test case is used to verify if include queryOptions.OptimizeLookups, must not change the contents of a successful protocol server response.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC42_GetListItemChangesSinceToken_OptimizeLookups()
        {
            // Add a list and add 3 list items.
            string listId = this.AddListToSUT(TemplateType.Generic_List);
            List<string> addedListitemIds = this.AddItemsToList(listId, 3);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Call GetListItemChangesSinceToken operation without "OptimizeLookups"
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null);

                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);
                this.Site.Assert.AreEqual<int>(
                                3,
                                zrowItems.Count(),
                                "The zrow items should equal to the number of added list items[3].");

                // Verify all added list item ids are in the zrow items' array.
                foreach (string addedListItemIdItem in addedListitemIds)
                {
                    bool isListItemIdIncludedInZrowItem = this.VerifyContainExpectedListItemById(addedListItemIdItem, zrowItems);
                    this.Site.Assert.IsTrue(
                                        isListItemIdIncludedInZrowItem,
                                        "The list item id[{0}] should be included in zrow items' array.",
                                        addedListItemIdItem);
                }

                // Setting "OptimizeLookups = true"
                CamlQueryOptions queryOptions = new CamlQueryOptions();
                queryOptions.QueryOptions = new CamlQueryOptionsQueryOptions();
                queryOptions.QueryOptions.OptimizeLookups = bool.TrueString;

                // Call GetListItemChangesSinceToken operation with the "OptimizeLookups = true"
                listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            null,
                                            queryOptions,
                                            null,
                                            null);

                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItemsWithOptimizeLookups = this.GetZrowItems(listItemChangesRes.listitems.data.Any);
                this.Site.Assert.AreEqual<int>(
                            3,
                            zrowItemsWithOptimizeLookups.Count(),
                            "The zrow items should equal to the number of added list items[3].");

                // Verify all added list item ids are in the zrow items' array.
                foreach (string addedListItemIdItem in addedListitemIds)
                {
                    bool isListItemIdIncludedInZrowItem = this.VerifyContainExpectedListItemById(addedListItemIdItem, zrowItems);
                    this.Site.Assert.IsTrue(
                                        isListItemIdIncludedInZrowItem,
                                        "The list item id[{0}] should be included in zrow items' array.",
                                        addedListItemIdItem);
                }

                // If pass upon verification, that means OptimizeLookups does not change the contents of a successful protocol server response. 
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1184
                this.Site.CaptureRequirement(
                                1184,
                                @"[In GetListItemChangesSinceToken]Including this element[queryOptions.OptimizeLookups]MUST NOT change the contents of a successful protocol server response.");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC43_DeleteDocumentsAndFolders
        /// <summary>
        /// This test case is used to verify if the folder item is deleted, then all documents and folders in it MUST be deleted too.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC43_DeleteDocumentsAndFolders()
        {
            this.Site.Assume.IsTrue(
                        Common.IsRequirementEnabled(106802, this.Site),
                        "Test is executed only when R106802Enabled is set to true.");

            // Add document library
            string listTitle = this.GetUniqueListName(TemplateType.Document_Library.ToString());
            string listId = this.AddListToSUT(listTitle, TemplateType.Document_Library);

            // Add a folder item into this document library.
            string folderName = this.GetUniqueFolderName();
            this.AddFolderIntoList(listId, folderName);

            // Upload a file into this folder
            string uploadFileName = this.GetUniqueUploadFileName();
            string fileUrl = SutControlAdapter.UploadFileWithFolder(listTitle, folderName, uploadFileName);
            this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(fileUrl),
                            "Uploading file to folder[{0}] under list[{1}] should succeed.",
                            folderName,
                            listTitle);

            // Call the GetListItemChangesSinceToken to get the added two items.
            // Set "<ViewFields />" in order to show all fields' value of a list.
            CamlViewFields viewfieds = new CamlViewFields();
            viewfieds.ViewFields = new CamlViewFieldsViewFields();

            // Call GetListItemChanges operation to get list items change.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
            listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                        listId,
                                        null,
                                        null,
                                        viewfieds,
                                        null,
                                        null,
                                        null,
                                        null);

            // Get the list items change data.
            this.VerifyContainZrowDataStructure(listItemChangesRes);
            XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);
            this.Site.Assert.IsNotNull(listItemChangesRes.listitems.Changes, "The Changes element should have value.");
            string tokenValue = listItemChangesRes.listitems.Changes.LastChangeToken;
            this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(tokenValue),
                                "The LastChangeToken should have value.");

            this.Site.Assert.AreEqual<int>(
                            2,
                            zrowItems.Count(),
                            "The response of GetListItemChangesSinceToken operation should contain two list items' change record.");

            // Verify these two item are contain in the response of GetListItemChangesSinceToken. If there are no any match zrow items found in zrow items array, the GetZrowItemIndexByFileRef method will throw Assert exception.
            this.GetZrowItemIndexByFileRef(zrowItems, folderName);
            this.GetZrowItemIndexByFileRef(zrowItems, uploadFileName);

            // Delete the folder item.
            bool deleteFolderResult = SutControlAdapter.DeleteFolder(listTitle, folderName);
            this.Site.Assert.IsTrue(
                        deleteFolderResult,
                        "Delete the folder[{0}] should succeed.",
                        folderName);

            listItemChangesRes = null;
            listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                        listId,
                                        null,
                                        null,
                                        viewfieds,
                                        null,
                                        null,
                                        tokenValue,
                                        null);

            // Get the list items change data after delete folder.
            this.VerifyContainZrowDataStructure(listItemChangesRes);
            this.Site.Assert.AreEqual<string>(
                "0",
                listItemChangesRes.listitems.data.ItemCount,
                "There should be no any list items return from protocol SUT after deleting the folder.");

            // Call HTTPGET method to try to get the upload file.
            HttpStatusCode statusCode = HttpStatusCode.OK;
            Uri fullUrlOfAttachmentPath;
            if (!Uri.TryCreate(fileUrl, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
            {
                this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
            }

            try
            {
                OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");
            }
            catch (WebException webEx)
            {
                statusCode = this.GetStatusCodeFromWebException(webEx);
            }

            // If the HTTPGET method return 404 NotFound exception, that means the upload file are deleted when deleted the folder.
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R66
            this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                                HttpStatusCode.NotFound,
                                statusCode,
                                66,
                                @"[In Documents]If the folder item is deleted, then all documents and folders in it MUST be deleted too.");
        }

        #endregion

        #region MSOUTSPS_S02_TC44_GenericList_VerifyVtiVersionHistoryValue
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC44_GenericList_VerifyVtiVersionHistoryValue()
        {
            string listId = this.AddListToSUT(TemplateType.Generic_List);

            // Add one item
            List<string> listItemIds = this.AddItemsToList(listId, 1);
            string listItemId = listItemIds[0];

            // Set "<ViewFields />" settings in order to view the properties' in properties bag, more detail is described in [MS-LISTSWS] section 2.2.4.5
            CamlViewFields viewfieds = new CamlViewFields();
            viewfieds.ViewFields = new CamlViewFieldsViewFields();
            viewfieds.ViewFields.Properties = bool.TrueString;
            viewfieds.ViewFields.FieldRef = new CamlViewFieldsViewFieldsFieldRef[1];
            viewfieds.ViewFields.FieldRef[0] = new CamlViewFieldsViewFieldsFieldRef();
            viewfieds.ViewFields.FieldRef[0].Name = "MetaInfo";

            XmlNode[] zrowItemsOfAddedItem = this.GetListItemsChangesFromSUT(listId, viewfieds);
            List<Dictionary<string, int>> vtiVersionHistoryValueOfAdded = this.GetVtiVersionHistoryValue(zrowItemsOfAddedItem, listItemId);

            // Update the existing item, make the vti_versionHistory change.
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.Update);

            // Set the target existing list item id
            Dictionary<string, string> updateSettingsOfListItem = new Dictionary<string, string>();
            updateSettingsOfListItem.Add("ID", listItemId);
            updateSettingsOfListItem.Add("Title", this.GetUniqueListItemTitle("Generic"));
            List<Dictionary<string, string>> updatedSettings = new List<Dictionary<string, string>>();
            updatedSettings.Add(updateSettingsOfListItem);
            UpdateListItemsUpdates updates = this.CreateUpdateListItems(cmds, updatedSettings, OnErrorEnum.Continue);

            // Update the existing list item.
            OutspsAdapter.UpdateListItems(listId, updates);

            XmlNode[] zrowItemsOfUpdatedItem = this.GetListItemsChangesFromSUT(listId, viewfieds);
            List<Dictionary<string, int>> vtiVersionHistoryValueOfUpdated = this.GetVtiVersionHistoryValue(zrowItemsOfUpdatedItem, listItemId);

            string guidPartValueOfAdded = vtiVersionHistoryValueOfAdded[0].ElementAt(0).Key;
            int integerPartValueOfAdded = vtiVersionHistoryValueOfAdded[0].ElementAt(0).Value;

            string guidPartValueOfUpdated = vtiVersionHistoryValueOfUpdated[0].ElementAt(0).Key;
            int integerPartValueOfUpdated = vtiVersionHistoryValueOfUpdated[0].ElementAt(0).Value;

            Guid guidOfAddedItem;
            string expectedGuidFormat = @"N";
            if (!Guid.TryParseExact(guidPartValueOfAdded, expectedGuidFormat, out guidOfAddedItem))
            {
                this.Site.Assert.Fail("The GUID part should be a valid GUID format, actual value:[{0}]", guidPartValueOfAdded);
            }

            Guid guidOfUpddatedItem;
            if (!Guid.TryParseExact(guidPartValueOfUpdated, expectedGuidFormat, out guidOfUpddatedItem))
            {
                this.Site.Assert.Fail("The GUID part should be a valid GUID format, actual value:[{0}]", guidOfUpddatedItem);
            }

            // If the GUID part value is it is "N" format ("hexadecimal string with no non-hexadecimal characters" format), then capture R218.
            this.Site.CaptureRequirement(
                                        218,
                                        @"[In Common Schema][vti_versionhistory:]GUIDs are written as a hexadecimal string with no non-hexadecimal characters.");

            // If the GUID part value is same GUID value, then capture R214, R229, R219
            bool isUseSameGuid = guidOfAddedItem.Equals(guidOfUpddatedItem);

            this.Site.Log.Add(
                            LogEntryKind.Debug,
                            @"GUID part value[{0}] in add item operation should be equal to the value[{0}] in update item.",
                            guidOfAddedItem,
                            guidOfUpddatedItem);

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R214
            this.Site.CaptureRequirementIfIsTrue(
                                           isUseSameGuid,
                                           214,
                                           @"[In Common Schema][vti_versionhistory:]Each unique GUID MUST appear at most once in the list.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R229
            this.Site.CaptureRequirementIfIsTrue(
                                           isUseSameGuid,
                                           229,
                                           @"[In Common Schema][vti_versionhistory:]This GUID[generated by every client and server that edits items] is reused and not regenerated.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R219
            this.Site.CaptureRequirementIfIsTrue(
                                           isUseSameGuid,
                                           219,
                                           @"[In Common Schema][vti_versionhistory:]Version history is built in the following way:Every protocol client and protocol server that edits items generates a GUID.");

            // If the integer part value is increased, then capture the R215, R222, R224
            bool isIntegerIncrease = integerPartValueOfUpdated > integerPartValueOfAdded;

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R215
            this.Site.CaptureRequirementIfIsTrue(
                                         isIntegerIncrease,
                                         215,
                                         @"[In Common Schema][vti_versionhistory:]One integer MUST be greater than all other integers.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R222
            this.Site.CaptureRequirementIfIsTrue(
                                        isIntegerIncrease,
                                        222,
                                        @"[In Common Schema][vti_versionhistory:]Each time an item is updated, the highest integer is found among the GUID-integer pairs.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R224
            this.Site.CaptureRequirementIfIsTrue(
                                        isIntegerIncrease,
                                        224,
                                        @"[In Common Schema][Version history is built in the following way: 2.]That integer is incremented by one to get the integer to use in step '3'.");

            // If the GUID part and Integer part value are pass the upon verification, then capture R212, R592
            this.Site.Assert.IsTrue(isIntegerIncrease, "The Integer part value of vti_versionhistory should be increased from previous value.");
            this.Site.Assert.IsTrue(isUseSameGuid, "The GUID part value of vti_versionhistory should be equal to previous value for same protocol client.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R212
            this.Site.CaptureRequirement(
                                        212,
                                        @"[In Common Schema]vti_versionhistory: Version history is a list of GUIDs and integers in the following format: GUID:Integer,GUID:Integer,GUID:Integer");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R592
            this.Site.CaptureRequirement(
                                        592,
                                        @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]vti_versionhistory<22>[ Field.ID:]None defined[Field.Type:]None defined (see description following the table).");

            // If the update item's vti_versionhistory use the same GUID and only have one GUID-Integer name-value pairs, then capture R225, R226
            int nameValuePairsItemCounter = vtiVersionHistoryValueOfUpdated[0].Count;

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R225
            this.Site.CaptureRequirementIfAreEqual<int>(
                                        1,
                                        nameValuePairsItemCounter,
                                        225,
                                        @"[In Common Schema][vti_versionhistory:]The editing protocol client or protocol server searches for its GUID in the GUID-integer pairs and removes the pair if it is found.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R226
            this.Site.CaptureRequirementIfAreEqual<int>(
                                        1,
                                        nameValuePairsItemCounter,
                                        226,
                                        @"[In Common Schema][Version history is built in the following way: 3.]Then the editing protocol client or protocol server inserts a new GUID-integer pair using its GUID and the integer from step '2'.");
        }

        #endregion

        #region MSOUTSPS_S02_TC45_GenericList_VerifyFieldsInCommonDefinition
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC45_GenericList_VerifyFieldsInCommonDefinition()
        {
            string listId = this.AddListToSUT(TemplateType.Generic_List);

            // Add one item
            List<string> listItemIds = this.AddItemsToList(listId, 2);
            string listItemIdOfFirstItem = listItemIds[0];
            string listItemIdOfSecondItem = listItemIds[1];

            // Get fields values of updated item.
            XmlNode[] zrowItemsOfAddedItem = this.GetListItemsChangesFromSUT(listId);
            int zrowItemIndexOfFirstItem = this.GetZrowItemIndexByListItemId(zrowItemsOfAddedItem, listItemIdOfFirstItem);
            string owshiddenversionValueOfFirstAddedItem = Common.GetZrowAttributeValue(zrowItemsOfAddedItem, zrowItemIndexOfFirstItem, "ows_owshiddenversion");
            string createdValueOfAddedItem = Common.GetZrowAttributeValue(zrowItemsOfAddedItem, zrowItemIndexOfFirstItem, "ows_created");

            // Sleep the specified seconds, in order to make the "created" file value is different from the "modified" value.
            int sleepValueOfMilliseSeconds = Common.GetConfigurationPropertyValue<int>("DelayBetweenAddItemAndUpdateItem", this.Site) * 1000;
            System.Threading.Thread.Sleep(sleepValueOfMilliseSeconds);

            // Update existing list item.
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.Update);

            // Set the target existing list item id
            Dictionary<string, string> updateSettingsOfListItem = new Dictionary<string, string>();
            updateSettingsOfListItem.Add("ID", listItemIdOfFirstItem);
            updateSettingsOfListItem.Add("Title", this.GetUniqueListItemTitle("Generic"));
            List<Dictionary<string, string>> updatedSettings = new List<Dictionary<string, string>>();
            updatedSettings.Add(updateSettingsOfListItem);
            UpdateListItemsUpdates updates = this.CreateUpdateListItems(cmds, updatedSettings, OnErrorEnum.Continue);

            // Update the existing list item.
            OutspsAdapter.UpdateListItems(listId, updates);

            // Get fields values of updated item.
            XmlNode[] zrowItemsOfUpdated = this.GetListItemsChangesFromSUT(listId);
            zrowItemIndexOfFirstItem = this.GetZrowItemIndexByListItemId(zrowItemsOfAddedItem, listItemIdOfFirstItem);
            string owshiddenversionValueOfUpdatedItem = Common.GetZrowAttributeValue(zrowItemsOfUpdated, zrowItemIndexOfFirstItem, "ows_owshiddenversion");
            string createdValueOfUpdatedItem = Common.GetZrowAttributeValue(zrowItemsOfUpdated, zrowItemIndexOfFirstItem, "ows_created");
            string listItemIdValueOfUpdatedItem = Common.GetZrowAttributeValue(zrowItemsOfUpdated, zrowItemIndexOfFirstItem, "ows_ID");
            string modifiedValueOfUpdatedItem = Common.GetZrowAttributeValue(zrowItemsOfUpdated, zrowItemIndexOfFirstItem, "ows_Modified");

            // If "created" fields are date time format value and the value of updated item is equal to the value of added item, then capture R194.
            DateTime createdDateOfAddedItem;
            if (!DateTime.TryParse(createdValueOfAddedItem, out createdDateOfAddedItem))
            {
                this.Site.Assert.Fail("The created field value should be valid DateTime format value for added item. actual:[{0}]", createdValueOfAddedItem);
            }

            DateTime createdDateOfUpdatedItem;
            if (!DateTime.TryParse(createdValueOfUpdatedItem, out createdDateOfUpdatedItem))
            {
                this.Site.Assert.Fail("The created field value should be valid DateTime format value for updated item. actual:[{0}]", createdValueOfAddedItem);
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R194
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                                                createdDateOfAddedItem,
                                                createdDateOfUpdatedItem,
                                                194,
                                                @"[In Common Schema]Created: The date and time the item was created.");

            // If the ID field is a integer value and the list item id of first item is different from the list item id of the second item, then capture R196
            int listItemIdIntValueOfUpdatedItem;
            if (!int.TryParse(listItemIdValueOfUpdatedItem, out listItemIdIntValueOfUpdatedItem))
            {
                this.Site.Assert.Fail(
                        "The ID field value should be valid integer format value for the updated item. Actual value[{0}]",
                        listItemIdValueOfUpdatedItem);
            }

            int zrowIndexOfSecondAddedItem = this.GetZrowItemIndexByListItemId(zrowItemsOfUpdated, listItemIdOfSecondItem);
            string listItemIdValueOfSecondAddedItem = Common.GetZrowAttributeValue(zrowItemsOfUpdated, zrowIndexOfSecondAddedItem, "ows_ID");
            int listItemIdIntValueOfSecondAddedItem;
            if (!int.TryParse(listItemIdValueOfSecondAddedItem, out listItemIdIntValueOfSecondAddedItem))
            {
                this.Site.Assert.Fail(
                       "The ID field value should be valid integer format value for the first added item. Actual value[{0}]",
                       listItemIdValueOfSecondAddedItem);
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R196
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                                            listItemIdIntValueOfUpdatedItem,
                                            listItemIdIntValueOfSecondAddedItem,
                                            196,
                                            @"[In Common Schema]ID: An integer that uniquely identifies this item from all other items in the list.");

            // If Modified field value is valid DataTime, and the Modified is larger than the created field value then capture R196.
            DateTime modifiedDateOfUpdatedItem;
            if (!DateTime.TryParse(modifiedValueOfUpdatedItem, out modifiedDateOfUpdatedItem))
            {
                this.Site.Assert.Fail("The modified field value should be valid DateTime format value for updated item. actual:[{0}]", createdValueOfAddedItem);
            }

            this.Site.Assert.IsTrue(modifiedDateOfUpdatedItem > createdDateOfUpdatedItem, "The modified field value should larger than the created field value.");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R201
            this.Site.CaptureRequirement(
                                     201,
                                     @"[In Common Schema]Modified: Date and time that the item was last modified.");

            // If the owshiddenversion is integer format value and the value in updated item is increased from the value in added item.
            int owshiddenversionIntValueOfFirstAddedItem;
            if (!int.TryParse(owshiddenversionValueOfFirstAddedItem, out owshiddenversionIntValueOfFirstAddedItem))
            {
                this.Site.Assert.Fail("The owshiddenversion field value should be valid integer format value for first added item. actual:[{0}]", owshiddenversionValueOfFirstAddedItem);
            }

            int owshiddenversionIntValueOfUpdatedItem;
            if (!int.TryParse(owshiddenversionValueOfUpdatedItem, out owshiddenversionIntValueOfUpdatedItem))
            {
                this.Site.Assert.Fail("The owshiddenversion field value should be valid integer format value for updated item. actual:[{0}]", owshiddenversionValueOfUpdatedItem);
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R454
            this.Site.CaptureRequirementIfAreEqual<int>(
                                                owshiddenversionIntValueOfUpdatedItem,
                                                owshiddenversionIntValueOfFirstAddedItem + 1,
                                                454,
                                                @"[In Common Schema]owshiddenversion: This is an integer that increases by 1 for a sample of N(default N=2) times the item is modified on the protocol server.");
        }

        #endregion

        #region MSOUTSPS_S02_TC46_OperateOnListItems_VerifyContentTypeId
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC46_OperateOnListItems_VerifyContentTypeId()
        {
            // Add a list and add a list item into the list.
            string listIdOfAppointment = this.AddListToSUT(TemplateType.Events);
            List<string> listitemIdsOfAppointment = this.AddItemsToList(listIdOfAppointment, 1);

            // If the content id begin with expected value, then capture R185
            bool isVerifyR185 = this.VerifyContentTypeIdForSpecifiedList(listIdOfAppointment, listitemIdsOfAppointment[0], "0x0102");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R185
            this.Site.CaptureRequirementIfIsTrue(
                                  isVerifyR185,
                                  185,
                                  @"[In Common Schema]ContentTypeId begins with 0x0102 corresponds to Appointment Content / item type name.");

            // Add a list and add a list item into the list.
            string listIdOfContacts = this.AddListToSUT(TemplateType.Contacts);
            List<string> listitemIdsOfContacts = this.AddItemsToList(listIdOfContacts, 1);

            // If the content id begin with expected value, then capture R186
            bool isVerifyR186 = this.VerifyContentTypeIdForSpecifiedList(listIdOfContacts, listitemIdsOfContacts[0], "0x0106");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R186
            this.Site.CaptureRequirementIfIsTrue(
                                  isVerifyR186,
                                  186,
                                  @"[In Common Schema]ContentTypeId begins with 0x0106 corresponds to Contact Content / item type name.");

            // Add a list and add a list item into the list.
            string documentLibraryTitle = this.GetUniqueListName(TemplateType.Document_Library.ToString());
            string listIdOfDocLibrary = this.AddListToSUT(documentLibraryTitle, TemplateType.Document_Library);
            string fileNameOnSut = this.GetUniqueUploadFileName();
            string uploadFilePath = SutControlAdapter.AddOneFileToDocumentLibrary(documentLibraryTitle, fileNameOnSut);
            this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(uploadFilePath),
                            "The uploading file to list[{0}] process should be succeed. Expected file name:[{1}]",
                             documentLibraryTitle,
                             fileNameOnSut);

            // Current document library only have one uploaded file, so the list item id is 1.  
            bool isVerifyR189 = this.VerifyContentTypeIdForSpecifiedList(listIdOfDocLibrary, "1", "0x0101");

            // If the content id begin with expected value, then verify MS-OUTSPS requirement: MS-OUTSPS_R189
            this.Site.CaptureRequirementIfIsTrue(
                                  isVerifyR189,
                                  189,
                                  @"[In Common Schema]ContentTypeId begins with 0x0101 corresponds to Document Content / item type name.");

            // Add a folder into the document library.
            string folderName = this.GetUniqueFolderName();
            string listItemIdOfFolder = this.AddFolderIntoList(listIdOfDocLibrary, folderName);

            // The folder item could only get by GetListItemChangesSinceToken operation.
            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesResOfSinceToken = OutspsAdapter.GetListItemChangesSinceToken(
                                                        listIdOfDocLibrary,
                                                        null,
                                                        null,
                                                        viewfieds,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                this.VerifyContainZrowDataStructure(listItemChangesResOfSinceToken);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesResOfSinceToken.listitems.data.Any);
                int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, listItemIdOfFolder);
                string contentTypeId = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_ContentTypeId");

                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(contentTypeId),
                                @"The contentTypeId field should have value.");

                // If the ContentTypeId field value begin with 0x0120, then capture R190
                bool isBeginWithExpectedValue = contentTypeId.StartsWith("0x0120", StringComparison.OrdinalIgnoreCase);
                this.Site.Assert.IsTrue(
                                      isBeginWithExpectedValue,
                                      "The contentTypeId should begin with:[{0}], the actual value:[{1}]",
                                      "0x0120",
                                      contentTypeId);

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R190
                this.Site.CaptureRequirement(
                                      190,
                                      @"[In Common Schema]ContentTypeId begins with 0x0120 corresponds to Folder Content / item type name.");
            }

            // Add a list and add a list item into the list.
            string listIdOfTasks = this.AddListToSUT(TemplateType.Tasks);
            this.AddItemsToList(listIdOfTasks, 1);

            // If the content id begin with expected value, then capture R186
            bool isVerifyR191 = this.VerifyContentTypeIdForSpecifiedList(listIdOfTasks, listitemIdsOfContacts[0], "0x0108");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R191
            this.Site.CaptureRequirementIfIsTrue(
                                  isVerifyR191,
                                  191,
                                  @"[In Common Schema]ContentTypeId Begins With 0x0108 corresponds to Task Content / item type name.");
        }

        #endregion

        #region MSOUTSPS_S02_TC47_OperateOnListItems_VerifyEventTypeInterpretedAsZero
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC47_OperateOnListItems_VerifyEventTypeInterpretedAsZero()
        {
            string listId = this.AddListToSUT(TemplateType.Events);

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.New);

            // Setting appointment fields' setting
            DateTime eventDate = DateTime.Today.Date.AddDays(1);
            DateTime endDate = eventDate.AddHours(1);
            string eventDateValue = this.GetUTCFormatTimeString(eventDate);
            string endDateValue = this.GetUTCFormatTimeString(endDate);
            string eventTitle = this.GetUniqueListItemTitle("SingleAppoinment");

            Dictionary<string, string> settingsOfFields = new Dictionary<string, string>();
            settingsOfFields.Add("EventDate", eventDateValue);
            settingsOfFields.Add("EndDate", endDateValue);
            settingsOfFields.Add("EventType", "2");
            settingsOfFields.Add("Title", eventTitle);

            List<Dictionary<string, string>> addeditemsOfRecurrence = new List<Dictionary<string, string>>();
            addeditemsOfRecurrence.Add(settingsOfFields);

            // Add a single appointment item
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addeditemsOfRecurrence, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);
            XmlNode[] zrowItems = this.GetZrowItems(updateResult.Results[0].Any);

            // Verify the eventType field value whether match the definition.
            this.VerifySimpleTypeSchema(zrowItems);
            string eventTypeFieldValue = Common.GetZrowAttributeValue(zrowItems, 0, "ows_EventType");

            // If the eventType field value does not equal to 1, 3 or 4, that means the protocol SUT does not interpreted the value "2" as "1", "3", "4" values which mean recurrence appointment item, and left possible values of eventType are "2", "0", all of these values are means single appointment item. If meet this condition, then capture R1010. 
            if (eventTypeFieldValue.Equals("1") || eventTypeFieldValue.Equals("3") || eventTypeFieldValue.Equals("4"))
            {
                this.Site.Assert.Fail(
                    @"The protocol SUT should interpreted the eventType field value '2' as meaning of single appointment item, when protocol client set the value to '2'. actual interpreted value:[{0}]",
                    eventTypeFieldValue);
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1010
            this.Site.CaptureRequirement(
                                    1010,
                                    @"[In EventType][The enumeration value]2[of the type EventType means]MUST be interpreted as 0.");
        }

        #endregion

        #region MSOUTSPS_S02_TC48_OperationListItemsForContacts
        /// <summary>
        /// This test case is used to verify Contacts list's fields when updating a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC48_OperationListItemsForContacts()
        {
            string listId = this.AddListToSUT(TemplateType.Contacts);

            Dictionary<string, string> fieldsValueSettings = new Dictionary<string, string>();
            string phoneValue = string.Format("135{0}", this.GenerateRandomNumber(0, 99999999));
            string emailAddress = string.Format(@"{0}@{1}.com", this.GenerateRandomString(5), this.GenerateRandomString(5));
            string uriString = string.Format(@"www.{0}.com", this.GenerateRandomString(5));
            string uriPathValue = string.Format(@"{0}.htm", this.GenerateRandomString(5));
            UriBuilder randomUri = new UriBuilder("http", uriString, int.Parse(this.GenerateRandomNumber(80, 255)), uriPathValue);
            string urlValue = randomUri.ToString();

            fieldsValueSettings.Add("CellPhone", phoneValue);

            fieldsValueSettings.Add("Comments", this.GenerateRandomString(10));

            fieldsValueSettings.Add("Company", this.GenerateRandomString(10));

            fieldsValueSettings.Add("CompanyPhonetic", phoneValue);

            fieldsValueSettings.Add("Email", emailAddress);

            fieldsValueSettings.Add("FirstName", this.GenerateRandomString(10));

            fieldsValueSettings.Add("FirstNamePhonetic", this.GenerateRandomString(10));

            fieldsValueSettings.Add("FullName", this.GenerateRandomString(10));

            fieldsValueSettings.Add("HomePhone", phoneValue);

            fieldsValueSettings.Add("JobTitle", this.GenerateRandomString(10));

            fieldsValueSettings.Add("LastNamePhonetic", this.GenerateRandomString(10));

            fieldsValueSettings.Add("Title", this.GenerateRandomString(10));

            fieldsValueSettings.Add("WebPage", urlValue);

            fieldsValueSettings.Add("WorkAddress", this.GenerateRandomString(10));

            fieldsValueSettings.Add("WorkCity", this.GenerateRandomString(10));

            fieldsValueSettings.Add("WorkCountry", this.GenerateRandomString(10));

            fieldsValueSettings.Add("WorkFax", phoneValue);

            fieldsValueSettings.Add("WorkPhone", phoneValue);

            fieldsValueSettings.Add("WorkState", this.GenerateRandomString(10));

            fieldsValueSettings.Add("WorkZip", this.GenerateRandomString(10));

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            List<Dictionary<string, string>> addeditemsOfContactItem = new List<Dictionary<string, string>>();
            addeditemsOfContactItem.Add(fieldsValueSettings);

            // Add the contact item into the list.
            UpdateListItemsUpdates updatesOfContactItem = this.CreateUpdateListItems(cmds, addeditemsOfContactItem, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfContactItem);

            // Get list item id from the response of UpdateListItems operation.
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            // Get list items' changes
            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);

            // Ignore the webPage field, because the return value in response of GetListItemChangesSinceToken operation contain other contents. So the value in response is not fully match the value test suite set. 
            fieldsValueSettings.Remove("WebPage");
            this.VerifyFieldsValuesEqualToExpected(fieldsValueSettings, zrowItems[zrowIndex]);
            string webPageUrlValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_WebPage");
            this.Site.Assert.IsTrue(
                                    webPageUrlValue.IndexOf(urlValue, StringComparison.OrdinalIgnoreCase) >= 0,
                                    @"The WebPage field's actual value[{0}] should contain the sub string[{1}].",
                                    webPageUrlValue,
                                    urlValue);

            // Verify the Editor field whether contain the current user name.
            string editorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_Editor");
            string currentUserValue = this.GetUserTypeValue();
            this.VerifyUserType(editorValue, currentUserValue, "Editor");
        }

        #endregion

        #region MSOUTSPS_S02_TC49_OperationListItemsForDiscussion
        /// <summary>
        /// This test case is used to verify DiscussionBoard list's fields when updating a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC49_OperationListItemsForDiscussion()
        {
            string listId = this.AddListToSUT(TemplateType.Discussion_Board);

            byte[] messageDataOfFirstItem = this.GetMessageDataForAddDiscussionBoardItem();
            OutspsAdapter.AddDiscussionBoardItem(listId, messageDataOfFirstItem);

            // add other DiscussionBoard item. 
            byte[] messageDataOfSecondItem = this.GetMessageDataForAddDiscussionBoardItem();
            OutspsAdapter.AddDiscussionBoardItem(listId, messageDataOfSecondItem);

            Dictionary<string, string> fieldsValueSettings = new Dictionary<string, string>();

            fieldsValueSettings.Add("Body", this.GenerateRandomString(10));

            fieldsValueSettings.Add("Title", this.GenerateRandomString(10));

            // Setting the fields setting to update the added discussion board item.
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.Update);

            // Current test suite have two items, now only update the first item.
            fieldsValueSettings.Add("ID", "1");

            // Setting recurring setting
            List<Dictionary<string, string>> updateditemsOfDiscussionBoard = new List<Dictionary<string, string>>();
            updateditemsOfDiscussionBoard.Add(fieldsValueSettings);

            // Update the discussion board item into the list.
            UpdateListItemsUpdates updatesOfDiscussionBoardItem = this.CreateUpdateListItems(cmds, updateditemsOfDiscussionBoard, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfDiscussionBoardItem);

            // Get list item id from the response of UpdateListItems operation.
            this.VerifyResponseOfUpdateListItem(updateResult);
            List<string> updatedDiscussionBoardItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                        listId,
                                                        null,
                                                        null,
                                                        viewfieds,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);
                int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, updatedDiscussionBoardItemIds[0]);

                // Ignore the fieldsValueSettings
                string bodyValueInSetting = fieldsValueSettings["Body"];
                fieldsValueSettings.Remove("Body");
                this.VerifyFieldsValuesEqualToExpected(fieldsValueSettings, zrowItems[zrowIndex]);

                string bodyValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_Body");
                this.Site.Assert.IsTrue(
                                       bodyValue.IndexOf(bodyValueInSetting, StringComparison.OrdinalIgnoreCase) >= 0,
                                       @"The Body field's actual value[{0}] should contain the sub string[{1}].",
                                       bodyValue,
                                       bodyValueInSetting);

                string currentUserValue = this.GetUserTypeValue();
                string editorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_Editor");
                this.VerifyUserType(editorValue, currentUserValue, "Editor");

                string authorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_Author");
                this.VerifyUserType(authorValue, currentUserValue, "Author");

                // Current list have only two items.
                this.Site.Assert.AreEqual<int>(
                                     2,
                                     zrowItems.Length,
                                     "The current list identified by list id[{0}] should contain two items.",
                                     listId);

                string actualThreadIndexValueOfFirstItem = Common.GetZrowAttributeValue(zrowItems, 0, "ows_ThreadIndex");
                string actualThreadIndexValueOfSecondItem = Common.GetZrowAttributeValue(zrowItems, 1, "ows_ThreadIndex");

                // If the ThreadIndex values are different between the first item and the second item, then capture R720.
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R720
                this.Site.CaptureRequirementIfAreNotEqual<string>(
                                        actualThreadIndexValueOfFirstItem.ToLower(),
                                        actualThreadIndexValueOfSecondItem.ToLower(),
                                        720,
                                        @"[In Discussion-Specific Schema]ThreadIndex: A thread index string that uniquely identifies each discussion thread in a list.");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC50_OperationListItemsForTasks
        /// <summary>
        /// This test case is used to verify Tasks list's fields when updating a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC50_OperationListItemsForTasks()
        {
            string listId = this.AddListToSUT(TemplateType.Tasks);

            Dictionary<string, string> fieldsValueSettings = new Dictionary<string, string>();

            DateTime startDate = DateTime.Now.Date;
            DateTime dueDate = DateTime.Now.Date.AddDays(1);

            string startDateValue = this.GetUTCFormatTimeString(startDate);
            string dueDateValue = this.GetUTCFormatTimeString(dueDate);
            double percentValue = double.Parse(this.GenerateRandomNumber(1, 100)) / 100;
            string userTypeValue = this.GetUserTypeValue();

            fieldsValueSettings.Add("AssignedTo", userTypeValue);

            fieldsValueSettings.Add("Body", this.GenerateRandomString(10));

            fieldsValueSettings.Add("DueDate", dueDateValue);

            fieldsValueSettings.Add("PercentComplete", percentValue.ToString());

            fieldsValueSettings.Add("Priority", "1");

            fieldsValueSettings.Add("StartDate", startDateValue);

            fieldsValueSettings.Add("Status", "1");

            fieldsValueSettings.Add("Title", this.GenerateRandomString(10));

            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);

            // Setting recurring setting
            List<Dictionary<string, string>> addeditemsOfContactItem = new List<Dictionary<string, string>>();
            addeditemsOfContactItem.Add(fieldsValueSettings);

            // Add the contact item into the list.
            UpdateListItemsUpdates updatesOfContactItem = this.CreateUpdateListItems(cmds, addeditemsOfContactItem, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfContactItem);

            // Get list item id from the response of UpdateListItems operation.
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);

            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);

            // Ignore the StartDate and DueDate, this test suite only verify whether the StartDate and DueDate's Date value whether equal to the setting value.
            fieldsValueSettings.Remove("StartDate");
            fieldsValueSettings.Remove("DueDate");

            // Ignore the PercentComplete field to perform fully match check, because the string value in response have different length from the value test suite set in request.
            fieldsValueSettings.Remove("PercentComplete");

            // Ignore the AssignedTo field to perform fully match check
            fieldsValueSettings.Remove("AssignedTo");

            this.VerifyFieldsValuesEqualToExpected(fieldsValueSettings, zrowItems[zrowIndex]);

            // Verify PercentComplete field
            string actualPercentCompleteValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_PercentComplete");
            double actualPercentComplete;
            if (!double.TryParse(actualPercentCompleteValue, out actualPercentComplete))
            {
                this.Site.Assert.Fail("The PercentComplete field value should be valid percent number.");
            }

            this.Site.Assert.AreEqual<double>(
                      percentValue,
                      actualPercentComplete,
                      "The PercentComplete field value in response of GetListItemChangesSinceToken should equal to expected set value.");

            // Verify StartDate and DueDate fields' value.
            string actualStartDateValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_StartDate");
            string actualDueDateValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_DueDate");
            this.VerifyDateTimeFieldValue(actualStartDateValue);
            this.VerifyDateTimeFieldValue(actualDueDateValue);

            string currentUserValue = this.GetUserTypeValue();

            // Verify the Editor field whether contain the current user name.
            string editorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_Editor");
            this.VerifyUserType(editorValue, currentUserValue, "Editor");

            // Verify the AssignedTo field
            string assignedToVlue = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_AssignedTo");
            this.VerifyUserType(assignedToVlue, currentUserValue, "AssignedTo");
        }

        #endregion

        #region MSOUTSPS_S02_TC51_OperationListItemsForDocument
        /// <summary>
        /// This test case is used to verify Document list's fields when updating a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC51_OperationListItemsForDocument()
        {
            // Add document library
            string documentLibraryTitle = this.GetUniqueListName(TemplateType.Document_Library.ToString());
            string listId = this.AddListToSUT(documentLibraryTitle, TemplateType.Document_Library);

            // Add a folder into the list.
            string folderItemName = this.GetUniqueFolderName();
            this.AddFolderIntoList(listId, folderItemName);

            // Upload a file to the folder.
            string fileName = this.GetUniqueUploadFileName();
            string fileUrl = SutControlAdapter.UploadFileWithFolder(documentLibraryTitle, folderItemName, fileName);
            this.Site.Assert.IsFalse(
                           string.IsNullOrEmpty(fileUrl),
                           "The uploading file to list[{0}] process should be succeed. Expected file name:[{1}]",
                            documentLibraryTitle,
                            fileName);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                                        listId,
                                                        null,
                                                        null,
                                                        viewfieds,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                // Get the list items' changes.
                this.VerifyContainZrowDataStructure(listItemChangesRes);
                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                // Current list only have two items, the first is the folder item, the second is the file item. so the list item is 2.
                int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, "2");

                string encodedAbsUrl = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_EncodedAbsUrl");
                string fileSizeDisplay = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_FileSizeDisplay");
                string linkFilename = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_LinkFilename");

                // If the upload file can be get by the HTTP GET method correctly by using encodedAbsUrl field value, then capture R728.
                this.Site.Assert.IsFalse(
                               string.IsNullOrEmpty(encodedAbsUrl),
                               "The encodedAbsUrl field should have value.");

                // Get the document item content.
                Uri fullUrlOfAttachmentPath;
                if (!Uri.TryCreate(encodedAbsUrl, UriKind.RelativeOrAbsolute, out fullUrlOfAttachmentPath))
                {
                    this.Site.Assert.Fail("The full url of attachment should be valid Uri format string.");
                }

                byte[] documentItemData = null;
                documentItemData = OutspsAdapter.HTTPGET(fullUrlOfAttachmentPath, "f");
                this.Site.Assert.IsNotNull(
                                documentItemData,
                                "The upload file[{0}] should be get by the HTTP GET method correctly by using encodedAbsUrl field value[{1}].",
                                fileName,
                                encodedAbsUrl);

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R728
                this.Site.CaptureRequirement(
                                728,
                                @"[In Document-Specific Schema]This[EncodedAbsUrl] MUST be present and valid.");

                // If the downloaded file content's size equal to the fileSizeDisplay field, then capture R734
                this.Site.Assert.IsFalse(
                              string.IsNullOrEmpty(fileSizeDisplay),
                              "The fileSizeDisplay field should have value.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R734
                this.Site.CaptureRequirementIfAreEqual<string>(
                                    fileSizeDisplay,
                                    documentItemData.Count().ToString(),
                                    734,
                                    "[In Document-Specific Schema]This[FileSizeDisplay] MUST be present and valid.");

                // If the linkFilename field value equal to the upload file name specified in previous step, then capture R741.
                this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(linkFilename),
                            "The linkFilename field should have value.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R741
                this.Site.CaptureRequirementIfAreEqual<string>(
                                    fileName.ToLower(),
                                    linkFilename.ToLower(),
                                    741,
                                    @"[In Document-Specific Schema]This[LinkFilename] MUST be present and valid.");

                if (Common.IsRequirementEnabled(732, this.Site))
                {
                    // Verify the fileDirRef field, it should contain the parent path of the upload file: rootFolder(documentLibraryTitle)/folder(folderItemName).
                    string fileDirRef = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_FileDirRef");
                    this.Site.Assert.IsFalse(
                             string.IsNullOrEmpty(fileDirRef),
                             "The fileDirRef field should have value.");

                    string actualFileDirRefValue = this.ParseRegularLookUpType(fileDirRef);
                    string expectedDirPath = string.Format(@"{0}/{1}", documentLibraryTitle, folderItemName);
                    this.Site.Assert.IsTrue(
                                        actualFileDirRefValue.IndexOf(expectedDirPath, StringComparison.OrdinalIgnoreCase) >= 0,
                                        "The FileDirRef field's value should include the expected path. FileDirRef field[{0}]\r\n Expected path[{1}] ",
                                        actualFileDirRefValue,
                                        expectedDirPath);

                    // Verify MS-OUTSPS requirement: MS-OUTSPS_R732
                    this.Site.CaptureRequirement(
                                        732,
                                        @"[In Appendix B: Product Behavior] Implementation does appear at the beginning of the FileDirRef value.(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                }

                // Verify user type fields for the folder item, the folder item is added as first item of the list.
                int zrowIndexOfFolderItem = this.GetZrowItemIndexByListItemId(zrowItems, "1");
                string currentUserValue = this.GetUserTypeValue();
                string authorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfFolderItem, "ows_Author");
                this.VerifyUserType(authorValue, currentUserValue, "Author");

                string editorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfFolderItem, "ows_Editor");
                this.VerifyUserType(editorValue, currentUserValue, "Editor");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC52_OperateOnListItems_VerifyMoreChangeValue
        /// <summary>
        /// This test case is used to verify MoreChange value when updating a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC52_OperateOnListItems_VerifyMoreChangeValue()
        {
            string listId = this.AddListToSUT(TemplateType.Generic_List);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Call GetListItemChangesSinceToken operation to get a token.
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                          listId,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null,
                                          null);

                if (null == listItemChangesRes || null == listItemChangesRes.listitems || null == listItemChangesRes.listitems.Changes)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken should contain the valid the changes data.");
                }

                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listItemChangesRes.listitems.Changes.LastChangeToken),
                                "The response of GetListItemChangesSinceToken should contain valid ChangeToken.");

                string changeTokenValue = listItemChangesRes.listitems.Changes.LastChangeToken;

                // Add 10 list items.
                this.AddItemsToList(listId, 10);

                // Setting "<viewFields />" to view all fields of the list.
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                // Set the low limited value equal to 2, apply the change token in the request, in order to the actual response contain more change flag.
                int rowLimitedValue = 2;
                listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId,
                                            null,
                                            null,
                                            viewfieds,
                                            rowLimitedValue.ToString(),
                                            null,
                                            changeTokenValue,
                                            null);

                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                this.Site.Assert.AreEqual<int>(
                                            rowLimitedValue,
                                            zrowItems.Length,
                                            "The return zrow item should match the specified row limit value.");

                if (null == listItemChangesRes || null == listItemChangesRes.listitems || null == listItemChangesRes.listitems.Changes)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken should contain the valid the changes data.");
                }

                // If the MoreChanges present and its value equal to TRUE, then capture R1225.
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listItemChangesRes.listitems.Changes.MoreChanges),
                                "The response of GetListItemChangesSinceToken should contain MoreChanges value.");

                string moreChangesValue = listItemChangesRes.listitems.Changes.MoreChanges;
                this.Site.Assert.IsTrue(
                                  moreChangesValue.Equals(bool.TrueString, StringComparison.OrdinalIgnoreCase),
                                  "The MoreChanges value should equal to 'TRUE'.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1225
                this.Site.CaptureRequirement(
                                    1225,
                                    @"[In GetListItemChangesSinceTokenResponse][If present condition meet][The attribute]Changes.MoreChanges is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");
            }
        }

        #endregion

        #region MSOUTSPS_S02_TC53_OperateOnListItems_VerifyListItemCollectionPositionNextValue
        /// <summary>
        /// This test case is used to verify Tasks template when update a list item.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S02_TC53_OperateOnListItems_VerifyListItemCollectionPositionNextValue()
        {
            string listId = this.AddListToSUT(TemplateType.Generic_List);

            // Add 10 list items.
            this.AddItemsToList(listId, 10);

            if (Common.IsRequirementEnabled(106802, this.Site))
            {
                // Setting "<viewFields />" to view all fields of the list.
                CamlViewFields viewfieds = new CamlViewFields();
                viewfieds.ViewFields = new CamlViewFieldsViewFields();

                // Set the low limited value equal to 2 without applying the change token in the request, in order to the actual response contain ListItemCollectionPositionNextValue value.
                int rowListmitedValue = 2;
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult listItemChangesRes = null;
                listItemChangesRes = OutspsAdapter.GetListItemChangesSinceToken(
                                            listId,
                                            null,
                                            null,
                                            viewfieds,
                                            rowListmitedValue.ToString(),
                                            null,
                                            null,
                                            null);

                XmlNode[] zrowItems = this.GetZrowItems(listItemChangesRes.listitems.data.Any);

                this.Site.Assert.AreEqual<int>(
                                            rowListmitedValue,
                                            zrowItems.Length,
                                            "The return zrow item should match the specified row limit value.");

                if (null == listItemChangesRes || null == listItemChangesRes.listitems || null == listItemChangesRes.listitems.Changes)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken should contain the valid the changes data.");
                }

                // If the ListItemCollectionPositionNextValue present then capture R1226.
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listItemChangesRes.listitems.data.ListItemCollectionPositionNext),
                                "The response of GetListItemChangesSinceToken should contain ListItemCollectionPositionNextValue value.");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1226
                this.Site.CaptureRequirement(
                                    1226,
                                    @"[In GetListItemChangesSinceTokenResponse][If present condition meet][The attribute]data.ListItemCollectionPositionNext is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");
            }
        }
        #endregion

        #endregion Test cases

        #region private methods

        /// <summary>
        /// A method used to get the vit_VersionHistory value from the zrow items data. If the vit_VersionHistory does not present in list schema definition, the method will try to get the value from properties bag.(metaInfo field) If there are any invalid format data for the vit_VersionHistory, this method will throw Assert exception.
        /// </summary>
        /// <param name="zrowItems">A parameter represents the zrow items array which contain the vit_VersionHistory data.</param>
        /// <param name="listItemId">A parameter represents the id of list item whose data present in row items array. This method will verify the zrow item's ows_Id field to find the match item.</param>
        /// <returns>A return value represents the name-value pairs of the vit_VersionHistory value for one list item. Each name-value pairs is "GUID-Integer" format and split by "," symbol in the vit_VersionHistory value.</returns>
        private List<Dictionary<string, int>> GetVtiVersionHistoryValue(XmlNode[] zrowItems, string listItemId)
        {
            // Get versionHistoryValue by using "ows_vti_versionhistory" field name, if the protocol SUT implement this field in list definition.
            string versionHistoryValue = string.Empty;
            string vtiHistoryVersionFieldName = string.Format("ows_{0}", "vti_versionhistory");

            // visit the zrow items array, and get the value of zrow item whose ows_id field equal to the specified list item id.
            int indexOfZrowItem = this.GetZrowItemIndexByListItemId(zrowItems, listItemId);
            versionHistoryValue = Common.GetZrowAttributeValue(zrowItems, indexOfZrowItem, vtiHistoryVersionFieldName);

            // If could not get the vti_versionhistory value from the list definition's format, then try to get it from properties bag. More detail is described in [MS-LISTSWS] section 3.1.4.24.2.1 
            if (string.IsNullOrEmpty(versionHistoryValue))
            {
                vtiHistoryVersionFieldName = string.Format("ows_MetaInfo_{0}", "vti_versionhistory");
                versionHistoryValue = Common.GetZrowAttributeValue(zrowItems, indexOfZrowItem, vtiHistoryVersionFieldName);
            }

            List<Dictionary<string, int>> vtiVersionHistoryValue = new List<Dictionary<string, int>>();

            if (string.IsNullOrEmpty(versionHistoryValue))
            {
                return vtiVersionHistoryValue;
            }

            string[] splitValues = versionHistoryValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string versionHistoryValueItem in splitValues)
            {
                string[] nameValuePairsOfversionHistory = versionHistoryValueItem.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                if (nameValuePairsOfversionHistory.Length != 2)
                {
                    this.Site.Assert.Fail("The vti_versionhistory value should be following format[GUID:Integer] and split by ':' symbol.");
                }

                Dictionary<string, int> nameValuePairItem = new Dictionary<string, int>();
                string guidPartValue = nameValuePairsOfversionHistory[0];

                // Get integer part value
                int integerPartValue;
                if (!int.TryParse(nameValuePairsOfversionHistory[1], out integerPartValue))
                {
                    this.Site.Assert.Fail("The integer part value should be integer format.");
                }

                nameValuePairItem.Add(guidPartValue, integerPartValue);
                vtiVersionHistoryValue.Add(nameValuePairItem);
            }

            return vtiVersionHistoryValue;
        }

        /// <summary>
        /// A method used to verify whether the ContentTypeId field's value begin with expected value.
        /// </summary>
        /// <param name="listId">A parameter represents the id of a list.</param>
        /// <param name="listItemId">A parameter represents the list item id under the list.</param>
        /// <param name="expectedBeginWithString">A parameter represents the expected value which the ContentTypeId field's value should begin with.</param>
        /// <returns>Return 'true' indicating the ContentTypeId field's value begin with expected value.</returns>
        private bool VerifyContentTypeIdForSpecifiedList(string listId, string listItemId, string expectedBeginWithString)
        {
            if (string.IsNullOrEmpty(listId))
            {
                throw new ArgumentNullException("listId");
            }

            if (string.IsNullOrEmpty(listItemId))
            {
                throw new ArgumentNullException("listItemId");
            }

            if (string.IsNullOrEmpty(expectedBeginWithString))
            {
                throw new ArgumentNullException("expectedBeginWithString");
            }

            // Get fields values of updated item.
            XmlNode[] zrowItems = this.GetListItemsChangesFromSUT(listId);
            int zrowIndex = this.GetZrowItemIndexByListItemId(zrowItems, listItemId);
            string contentTypeId = Common.GetZrowAttributeValue(zrowItems, zrowIndex, "ows_ContentTypeId");

            this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(contentTypeId),
                            @"The contentTypeId field should have value.");

            bool isBeginWithExpectedValue = contentTypeId.StartsWith(expectedBeginWithString, StringComparison.OrdinalIgnoreCase);
            this.Site.Assert.IsTrue(
                                  isBeginWithExpectedValue,
                                  "The contentTypeId should begin with:[{0}], the actual value:[{1}]",
                                  expectedBeginWithString,
                                  contentTypeId);

            return isBeginWithExpectedValue;
        }

        /// <summary>
        /// A method used to verify fields' values whether equal to the fields setting in request of UpdateListItems operation. If there are any verification errors, method will throw Assert exception.
        /// </summary>
        /// <param name="fieldsSettings">A parameter represents the fields setting in request of UpdateListItems operation.</param>
        /// <param name="zrowItem">A parameter represents the zrow item which contain a list item's fields' values.</param>
        /// <returns>Return 'true' indicating that fields' values whether equal to the fields setting in request of UpdateListItems operation.</returns>
        private bool VerifyFieldsValuesEqualToExpected(Dictionary<string, string> fieldsSettings, XmlNode zrowItem)
        {
            if (null == fieldsSettings)
            {
                throw new ArgumentNullException("fieldsSettings");
            }

            if (null == zrowItem)
            {
                throw new ArgumentNullException("zrowItem");
            }

            if (0 == fieldsSettings.Count)
            {
                throw new ArgumentException("The fieldsValues should contain at least one fields value record.", "fieldsSettings");
            }

            StringBuilder errorStrBuilder = new StringBuilder();
            foreach (string fieldNameItem in fieldsSettings.Keys)
            {
                string expectedValue = fieldsSettings[fieldNameItem];
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(expectedValue),
                                    @"The [{0}] field value should not be empty.",
                                    fieldNameItem);

                string actualFieldValue = string.Empty;
                string expectedFieldName = string.Format(@"ows_{0}", fieldNameItem);

                // Get the actual value for the field.
                var actualFieldValues = from XmlAttribute attributeItem in zrowItem.Attributes
                                        where attributeItem.Name.Equals(expectedFieldName, StringComparison.OrdinalIgnoreCase)
                                        select attributeItem.Value;

                if (actualFieldValues.Count() > 0)
                {
                    actualFieldValue = actualFieldValues.ElementAt<string>(0);
                    bool isActualValueEqualToExpected = actualFieldValue.Equals(expectedValue, StringComparison.OrdinalIgnoreCase);

                    if (!isActualValueEqualToExpected)
                    {
                        errorStrBuilder.Append(string.Format("The [{0}] field's actual value[{1}] should equal to expected value[{2}].\r\n", fieldNameItem, actualFieldValue, expectedValue));
                    }
                }
                else
                {
                    errorStrBuilder.Append(string.Format("The [{0}] field's value could not be found in the zrow item.\r\n", fieldNameItem));
                }
            }

            if (errorStrBuilder.Length != 0)
            {
                this.Site.Assert.Fail("There are some verification errors:\r\n{0}", errorStrBuilder.ToString());
            }

            return true;
        }

        /// <summary>
        /// A method used to verify DateTime value of a field whether equal to the value set in the request. If there are any errors this method will throw Assert exception.
        /// </summary>
        /// <param name="actualdateTimeValue">A parameter represents the actual DateTime value of string format, which is get from zrow items array.</param>
        private void VerifyDateTimeFieldValue(string actualdateTimeValue)
        {
            DateTime actualDateTime;
            if (!DateTime.TryParse(actualdateTimeValue, null, System.Globalization.DateTimeStyles.AdjustToUniversal, out actualDateTime))
            {
                this.Site.Assert.Fail("The actual DateTime value is not a valid DateTime format. Actual value:[{0}]", actualdateTimeValue);
            }
        }

        /// <summary>
        /// A method used to get a valid value for a user type field. This method will add a Generic list and add a list item, and then get the list item and parse the "Editor" field value.
        /// </summary>
        /// <returns>A return value represents the valid user value for the current user the test suite use. This value is only available on the first time setting the user type value in a list. If need to compare the value of a list item, must note how many times the list item have been updated, the identifier of this value is increased for each update operation for a list item.</returns>
        private string GetUserTypeValue()
        {
            string listTitle = this.GetUniqueListName("GetCurrentUserParticipants");
            string listId = this.AddListToSUT(listTitle, TemplateType.Generic_List);

            // set the field value for new item.
            Dictionary<string, string> fieldsSetting = new Dictionary<string, string>();
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = this.GenerateRandomString(5);
            fieldsSetting.Add(fieldName, fieldValue);

            List<Dictionary<string, string>> addItemsSetting = new List<Dictionary<string, string>>();
            addItemsSetting.Add(fieldsSetting);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.New);

            // Add list item.
            UpdateListItemsUpdates updatesOfRecurrence = this.CreateUpdateListItems(cmds, addItemsSetting, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult updateResult = OutspsAdapter.UpdateListItems(listId, updatesOfRecurrence);

            // Get the listItem id.
            this.VerifyResponseOfUpdateListItem(updateResult);
            XmlNode[] zrowItems = this.GetZrowItems(updateResult.Results[0].Any);
            List<string> addedRecurrenceItemIds = this.GetListItemIdsFromUpdateListItemsResponse(updateResult, 1);
            int zrowIndexOfaddedItem = this.GetZrowItemIndexByListItemId(zrowItems, addedRecurrenceItemIds[0]);
            string actualEditorValue = Common.GetZrowAttributeValue(zrowItems, zrowIndexOfaddedItem, "ows_Editor");

            this.Site.Assert.IsFalse(
                            string.IsNullOrEmpty(actualEditorValue),
                            "The Editor field value should have value.");

            string[] userTypeDataTemp = actualEditorValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            var validUserValues = from userValueItem in userTypeDataTemp
                                  where userValueItem.IndexOf(";#", StringComparison.OrdinalIgnoreCase) > 0
                                  select userValueItem;

            this.Site.Assert.AreNotEqual<int>(
                                 0,
                                 validUserValues.Count(),
                                 "The Editor field value should contain at least one valid user value which is split by ';#'.");

            string validUserValue = validUserValues.ElementAt<string>(0);
            return validUserValue;
        }

        /// <summary>
        /// A method used to verify whether a user type field contain the valid value of current user which the test suite use. If there are any verification errors, this method will throw Assert exception.
        /// </summary>
        /// <param name="actualUserTypeValue">A parameter represents the value of user type field.</param>
        /// <param name="validUserTypeValue">A parameter represents the valid value of current user which the test suite use. This value could be get by calling GetUserTypeValue method.</param>
        /// <param name="fieldName">A parameter represents the name of a field whose value is performed the verification.</param>
        private void VerifyUserType(string actualUserTypeValue, string validUserTypeValue, string fieldName)
        {
            if (string.IsNullOrEmpty(actualUserTypeValue))
            {
                throw new ArgumentNullException("fieldName");
            }

            if (string.IsNullOrEmpty(actualUserTypeValue))
            {
                throw new ArgumentNullException("validUserTypeValue");
            }

            this.Site.Assert.IsFalse(
                                   string.IsNullOrEmpty(actualUserTypeValue),
                                   "The [{0}] field value should have value.",
                                   fieldName);

            this.Site.Assert.IsTrue(
                                actualUserTypeValue.IndexOf(validUserTypeValue, StringComparison.OrdinalIgnoreCase) >= 0,
                                "The [{0}] field value should contain the current user name which the test suite use[{1}].",
                                fieldName,
                                validUserTypeValue);
        }

        /// <summary>
        /// A method used to parse regular LookUp type of a field. A LookUp type can be other format combinations, this method only parse the format 'identifier>;#value'.
        /// </summary>
        /// <param name="rawLookUpValue">A parameter represents the raw look up value from the response.</param>
        /// <returns>A return value represents the value contained in LookUp type string.</returns>
        private string ParseRegularLookUpType(string rawLookUpValue)
        {
            if (string.IsNullOrEmpty(rawLookUpValue))
            {
                throw new ArgumentNullException("rawLookUpValue");
            }

            // The fileDirRef field is split by ";#", more detail as descried in LookUp type of field.(section 2.3 in [MS-WSSTS])
            string[] fileDirRefValues = rawLookUpValue.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);
            this.Site.Assert.AreEqual<int>(
                                    2,
                                    fileDirRefValues.Length,
                                    "The fileDirRef value should be follow below format:[<identifier>;#<value>]");

            return fileDirRefValues[1];
        }

        #endregion private methods
    }
}