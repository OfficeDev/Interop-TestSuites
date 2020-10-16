namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is used to test the SmartForward command.
    /// </summary>
    [TestClass]
    public class S17_SmartForward : TestSuiteBase
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
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// This test case is used to verify the server returns an empty response, when mail forward successfully.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC01_SmartForward_Success()
        {
            #region Call SendMail command to send plain text email messages to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Call Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncChangeResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            string originalServerID = TestSuiteBase.FindServerId(syncChangeResponse, "Subject", emailSubject);
            string originalContent = TestSuiteBase.GetDataFromResponseBodyElement(syncChangeResponse, originalServerID);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call SmartForward command to forward messages without retrieving the full, original message from the server.
            string forwardSubject = string.Format("FW:{0}", emailSubject);
            SmartForwardRequest smartForwardRequest = this.CreateDefaultForwardRequest(originalServerID, forwardSubject, this.User2Information.InboxCollectionId);
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R568, MS-ASCMD_R4407
            // If the message was forwarded successfully, server returns an empty response without XML body, then MS-ASCMD_R568, MS-ASCMD_R4407 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R568");

            // Verify MS-ASCMD requirement: MS-ASCMD_R568
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                smartForwardResponse.ResponseDataXML,
                568,
                @"[In SmartForward] If the message was forwarded successfully, the server returns an empty response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4407");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4407
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                smartForwardResponse.ResponseDataXML,
                4407,
                @"[In Status(SmartForward and SmartReply)] If the SmartForward command request [or SmartReply command request] succeeds, no XML body is returned in the response.");
            #endregion

            #region After user2 forwarded email to user3, sync user3 mailbox changes
            this.SwitchUser(this.User3Information);
            SyncResponse syncForwardResult = this.GetMailItem(this.User3Information.InboxCollectionId, forwardSubject);
            string forwardItemServerID = TestSuiteBase.FindServerId(syncForwardResult, "Subject", forwardSubject);
            string forwardItemContent = TestSuiteBase.GetDataFromResponseBodyElement(syncForwardResult, forwardItemServerID);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User3Information.InboxCollectionId, forwardSubject);

            #endregion

            // Compare original content with forward content
            bool isContained = forwardItemContent.Contains(originalContent);

            #region Verify Requirements MS-ASCMD_R543, MS-ASCMD_R532
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R543");

            // Verify MS-ASCMD requirement: MS-ASCMD_R543
            Site.CaptureRequirementIfIsTrue(
                isContained,
                543,
                @"[In SmartForward] When the SmartForward command is used for a normal message or a meeting, the behavior of the SmartForward command is the same as that of the SmartReply command (section 2.2.2.18).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R532");

            // Verify MS-ASCMD requirement: MS-ASCMD_R532
            Site.CaptureRequirementIfIsNotNull(
                forwardItemServerID,
                532,
                @"[In SmartForward] The SmartForward command is used by clients to forward messages without retrieving the full, original message from the server.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify server returns status code, when SmartForward is failed.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC02_SmartForward_Fail()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The AccountID element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            #region Call SendMail command to send one plain text email messages to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Call Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncChangeResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            string originalServerId = TestSuiteBase.FindServerId(syncChangeResponse, "Subject", emailSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call SmartForward command to forward messages without retrieving the full, original message from the server
            // Create invalid SmartForward request
            SmartForwardRequest smartForwardRequest = new SmartForwardRequest
            {
                RequestData = new Request.SmartForward
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source
                    {
                        FolderId = this.User2Information.InboxCollectionId,
                        ItemId = originalServerId
                    },
                    Mime = string.Empty,
                    AccountId = "InvalidValueAccountID"
                }
            };

            smartForwardRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, this.User2Information.InboxCollectionId
                },
                {
                    CmdParameterName.ItemId, "5:" + originalServerId
                }
            });

            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4408");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4408
            // If SmartForward operation failed, server will return Status element in response, then MS-ASCMD_4408 is verified.
            Site.CaptureRequirementIfIsNotNull(
                smartForwardResponse.ResponseData.Status,
                4408,
                @"[In Status(SmartForward and SmartReply)] If the SmartForward command request [or SmartReply command request] fails, the Status element contains a code that indicates the type of failure.");
        }

        /// <summary>
        /// This test case is used to verify when the SmartForward command is used for an appointment, the original message is included as an attachment in the outgoing message.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC03_SmartForwardAppointment()
        {
            #region User1 calls Sync command uploading one calendar item to create one appointment
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            Calendar calendar = new Calendar
                {
                    OrganizerEmail = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                    OrganizerName = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                    UID = Guid.NewGuid().ToString(),
                    Subject = meetingRequestSubject
                };

            this.SyncAddCalendar(calendar);

            // Calls Sync command to sync user1's calendar folder
            SyncResponse syncUser1CalendarFolder = this.GetMailItem(this.User1Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemId = TestSuiteBase.FindServerId(syncUser1CalendarFolder, "Subject", meetingRequestSubject);

            // Record items need to be cleaned up.
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User1 calls smartForward command to forward mail to user2
            string forwardFromUser = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardSubject = string.Format("FW:{0}", meetingRequestSubject);
            string forwardContent = Common.GenerateResourceName(Site, "forward:Appointment body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(this.User1Information.CalendarCollectionId, calendarItemId, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
            #endregion

            #region User2 calls Sync command to get the forward mail sent by user1
            this.SwitchUser(this.User2Information);
            SyncResponse user2MailboxChange = this.GetMailItem(this.User2Information.InboxCollectionId, forwardSubject);

            // Record items need to be cleaned up.
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, forwardSubject);
            string mailItemServerId = TestSuiteBase.FindServerId(user2MailboxChange, "Subject", forwardSubject);

            Response.Attachments attachments = (Response.Attachments)TestSuiteBase.GetElementValueFromSyncResponse(user2MailboxChange, mailItemServerId, Response.ItemsChoiceType8.Attachments);
            Site.Assert.AreEqual<int>(1, attachments.Items.Length, "Server should return one attachment, if SmartForward one appointment executes successfully.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R542");

            // Verify MS-ASCMD requirement: MS-ASCMD_R542
            Site.CaptureRequirementIfIsTrue(
                attachments.Items.Length == 1 && ((Response.AttachmentsAttachment)attachments.Items[0]).DisplayName.Contains(".eml"),
                542,
                @"[In SmartForward] When the SmartForward command is used for an appointment, the original message is included by the server as an attachment to the outgoing message.");
        }

        /// <summary>
        /// This test case is used to verify if the value of the InstanceId element is invalid, the server returns status value 104.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC04_SmartForwardWithInvalidInstanceId()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls SmartForward command to forward the calendar item to user3 with invalid InstanceId value in SmartForward request
            string forwardFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardSubject = string.Format("FW:{0}", meetingRequestSubject);
            string forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(this.User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

            // Set instanceID with format not the same as required format "2010-03-20T22:40:00.000Z".
            string instanceID = DateTime.Now.ToString();
            smartForwardRequest.RequestData.Source.InstanceId = instanceID;
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R541");

            // Verify MS-ASCMD requirement: MS-ASCMD_R541
            Site.CaptureRequirementIfAreEqual<string>(
                "104",
                smartForwardResponse.ResponseData.Status,
                541,
                @"[In SmartForward] If the value of the InstanceId element is invalid, the server responds with Status element (section 2.2.3.162.15) value 104, as specified in section 2.2.4.");
        }

        /// <summary>
        /// This test case is used to verify when SmartForward is applied to a recurring meeting, the InstanceId element specifies the ID of a particular occurrence in the recurring meeting.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC05_SmartForwardWithInstanceIdSuccess()
        {
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncMeetingMailResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);

            // Record relative items for clean up.
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to accept the meeting
            string serverIDForMeetingRequest = TestSuiteBase.FindServerId(syncMeetingMailResponse, "Subject", meetingRequestSubject);
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

            // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            Site.Assert.IsNotNull(meetingResponseResponse.ResponseData.Result[0].CalendarId, "If the meeting was accepted, server should return calendarId in response");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            this.GetMailItem(this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls Sync command to sync user calendar changes
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            string startTime = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
            Response.Recurrence recurrence = (Response.Recurrence)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.Recurrence);
            Site.Assert.IsNotNull(recurrence, "If user2 received recurring meeting request, the calendar item should contain recurrence element");

            // Record relative items for clean up.
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls SmartForward command to forward the calendar item to user3 with correct InstanceId value in SmartForward request
            string forwardFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardSubject = string.Format("FW:{0}", meetingRequestSubject);
            string forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(this.User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

            // Set instanceID with format the same as required format "2010-03-20T22:40:00.000Z".
            string instanceID = ConvertInstanceIdFormat(startTime);
            smartForwardRequest.RequestData.Source.InstanceId = instanceID;
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
            #endregion

            #region After user2 forwards email to user3, sync user3 mailbox changes
            this.SwitchUser(this.User3Information);
            SyncResponse syncForwardResult = this.GetMailItem(this.User3Information.InboxCollectionId, forwardSubject);
            string forwardItemServerID = TestSuiteBase.FindServerId(syncForwardResult, "Subject", forwardSubject);

            // Sync user3 Calendar folder 
            SyncResponse syncUser3CalendarFolder = this.GetMailItem(this.User3Information.CalendarCollectionId, forwardSubject);
            string user3CalendarItemID = TestSuiteBase.FindServerId(syncUser3CalendarFolder, "Subject", forwardSubject);

            // Record email items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User3Information.InboxCollectionId, forwardSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User3Information.CalendarCollectionId, forwardSubject);
            #endregion

            #region Record the meeting forward notification mail which sent from server to User1.
            this.SwitchUser(this.User1Information);
            string notificationSubject = "Meeting Forward Notification: " + forwardSubject;
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.DeletedItemsCollectionId, notificationSubject);
            this.GetMailItem(this.User1Information.DeletedItemsCollectionId, notificationSubject);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R538");

            // Verify MS-ASCMD requirement: MS-ASCMD_R538
            // If the calendar item with specified subject exists in user3 Calendar folder and email item exists in user3 Inbox folder which means user3 gets the forwarded mail.
            Site.CaptureRequirementIfIsTrue(
                user3CalendarItemID != null && forwardItemServerID != null,
                538,
                @"[In SmartForward] When SmartForward is applied to a recurring meeting, the InstanceId element (section 2.2.3.83.2) specifies the ID of a particular occurrence in the recurring meeting.");
        }

        /// <summary>
        /// This test case is used to verify when SmartForward request without the InstanceId element, the implementation forward the entire recurring meeting.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC06_SmartForwardRecurringMeetingWithoutInstanceId()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5834, this.Site), "[In Appendix A: Product Behavior] If SmartForward is applied to a recurring meeting and the InstanceId element is absent, the implementation does forward the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
            Site.Assume.AreNotEqual<string>("16.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.0");
            Site.Assume.AreNotEqual<string>("16.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "Recurrences cannot be added in protocol version 16.1");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);

            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);

            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            Response.Recurrence recurrence = (Response.Recurrence)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.Recurrence);
            Site.Assert.IsNotNull(recurrence, "If user2 received recurring meeting request, the calendar item should contain recurrence element");

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls SmartForward command to forward the calendar item to user3 without InstanceId element in SmartForward request
            string forwardFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardSubject = string.Format("FW:{0}", meetingRequestSubject);
            string forwardContent = Common.GenerateResourceName(Site, "forward:Meeting Instance body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(this.User2Information.CalendarCollectionId, calendarItemID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);

            smartForwardRequest.RequestData.Source.InstanceId = null;
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "If SmartForward command executes successfully, server should return empty xml data");
            #endregion

            #region After user2 forwards email to user3, sync user3 mailbox changes
            this.SwitchUser(this.User3Information);
            SyncResponse syncForwardResult = this.GetMailItem(this.User3Information.InboxCollectionId, forwardSubject);
            string forwardItemServerID = TestSuiteBase.FindServerId(syncForwardResult, "Subject", forwardSubject);
            Response.MeetingRequest user3Meeting = (Response.MeetingRequest)TestSuiteBase.GetElementValueFromSyncResponse(syncForwardResult, forwardItemServerID, Response.ItemsChoiceType8.MeetingRequest);

            // Record email items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User2Information.InboxCollectionId, forwardSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User2Information.CalendarCollectionId, forwardSubject);
            #endregion

            #region Check the meeting forward notification mail which is sent from server to User1.
            this.SwitchUser(this.User1Information);
            string notificationSubject = "Meeting Forward Notification: " + forwardSubject;
            this.CheckMeetingForwardNotification(this.User1Information, notificationSubject);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5834");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5834
            // If the calendar item with specified subject contains Recurrence element, which indicates user3 received the entire meeting request.
            Site.CaptureRequirementIfIsTrue(
                user3Meeting.Recurrences != null && forwardItemServerID != null,
                5834,
                @"[In Appendix A: Product Behavior] If SmartForward is applied to a recurring meeting and the InstanceId element is absent, the implementation does forward the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to verify when ReplaceMime is present in the request, the body or attachment is not included.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S17_TC07_SmartForward_ReplaceMime()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "ReplaceMime is not support when MS-ASProtocolVersion header is set to 12.1.MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region Call SendMail command to send plain text email messages to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string emailBody = Common.GenerateResourceName(Site, "NormalAttachment_Body");
            this.SendEmailWithAttachment(emailSubject, emailBody);
            #endregion

            #region Call Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncChangeResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            string originalServerID = TestSuiteBase.FindServerId(syncChangeResponse, "Subject", emailSubject);
            string originalContent = TestSuiteBase.GetDataFromResponseBodyElement(syncChangeResponse, originalServerID);
            Response.AttachmentsAttachment[] originalAttachments = this.GetEmailAttachments(syncChangeResponse, emailSubject);
            Site.Assert.IsTrue(originalAttachments != null && originalAttachments.Length == 1, "The email should contain a single attachment.");

            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call SmartForward command to forward messages with ReplaceMime.
            string forwardSubject = string.Format("FW:{0}", emailSubject);
            SmartForwardRequest smartForwardRequest = this.CreateDefaultForwardRequest(originalServerID, forwardSubject, this.User2Information.InboxCollectionId);
            smartForwardRequest.RequestData.ReplaceMime = string.Empty;
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            #endregion

            #region After user2 forwarded email to user3, sync user3 mailbox changes
            this.SwitchUser(this.User3Information);
            SyncResponse syncForwardResult = this.GetMailItem(this.User3Information.InboxCollectionId, forwardSubject);
            string forwardItemServerID = TestSuiteBase.FindServerId(syncForwardResult, "Subject", forwardSubject);
            string forwardItemContent = TestSuiteBase.GetDataFromResponseBodyElement(syncForwardResult, forwardItemServerID);
            Response.AttachmentsAttachment[] forwardAttachments = this.GetEmailAttachments(syncForwardResult, forwardSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are used in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, this.User3Information.InboxCollectionId, forwardSubject);
            #endregion

            // Compare original content with forward content
            bool isContained = forwardItemContent.Contains(originalContent);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3775");

            Site.Assert.IsNull(
                forwardAttachments,
                @"The attachment should not be returned");

            Site.CaptureRequirementIfIsFalse(
                isContained,
                3775,
                @"[In ReplaceMime] When the ReplaceMime element is present, the server MUST not include the body or attachments of the original message being forwarded.");
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Try to parse the no separator time string to DateTime
        /// </summary>
        /// <param name="time">The specified DateTime string</param>
        /// <returns>Return the DateTime with instanceId specified format</returns>
        private static string ConvertInstanceIdFormat(string time)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(time.Substring(0, 4));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(4, 2));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(6, 5));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(11, 2));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(13, 2));
            stringBuilder.Append(".000");
            stringBuilder.Append(time.Substring(15));
            return stringBuilder.ToString();
        }

        /// <summary>
        /// Set sync request application data with calendar value
        /// </summary>
        /// <param name="calendar">The calendar instance</param>
        /// <returns>The application data for sync request</returns>
        private static Request.SyncCollectionAddApplicationData SetApplicationDataFromCalendar(Calendar calendar)
        {
            Request.SyncCollectionAddApplicationData applicationData = new Request.SyncCollectionAddApplicationData();
            List<Request.ItemsChoiceType8> elementName = new List<Request.ItemsChoiceType8>();
            List<object> elementValue = new List<object>();

            // Set application data
            elementName.Add(Request.ItemsChoiceType8.Timezone);
            elementValue.Add(calendar.Timezone);

            elementName.Add(Request.ItemsChoiceType8.Subject);
            elementValue.Add(calendar.Subject);

            elementName.Add(Request.ItemsChoiceType8.Sensitivity);
            elementValue.Add(calendar.Sensitivity);

            elementName.Add(Request.ItemsChoiceType8.BusyStatus);
            elementValue.Add(calendar.BusyStatus);

            elementName.Add(Request.ItemsChoiceType8.AllDayEvent);
            elementValue.Add(calendar.AllDayEvent);

            applicationData.ItemsElementName = elementName.ToArray();
            applicationData.Items = elementValue.ToArray();
            return applicationData;
        }

        /// <summary>
        /// Create default SmartForward request to forward an item from user 2 to user 3.
        /// </summary>
        /// <param name="originalServerID">The item serverID</param>
        /// <param name="forwardSubject">The forward mail subject</param>
        /// <param name="senderCollectionId">The sender inbox collectionId</param>
        /// <returns>The SmartForward request</returns>
        private SmartForwardRequest CreateDefaultForwardRequest(string originalServerID, string forwardSubject, string senderCollectionId)
        {
            string forwardFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User3Information.UserName, this.User3Information.UserDomain);
            string forwardContent = Common.GenerateResourceName(Site, "forward:body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(senderCollectionId, originalServerID, forwardFromUser, forwardToUser, string.Empty, string.Empty, forwardSubject, forwardContent);
            return smartForwardRequest;
        }

        /// <summary>
        /// Add a meeting or appointment to server
        /// </summary>
        /// <param name="calendar">the calendar item</param>
        private void SyncAddCalendar(Calendar calendar)
        {
            Request.SyncCollectionAddApplicationData applicationData = SetApplicationDataFromCalendar(calendar);

            this.GetInitialSyncResponse(this.User1Information.CalendarCollectionId);
            Request.SyncCollectionAdd addCalendar = new Request.SyncCollectionAdd
            {
                ClientId = TestSuiteBase.ClientId,
                ApplicationData = applicationData
            };

            SyncRequest syncAddCalendarRequest = TestSuiteBase.CreateSyncAddRequest(this.LastSyncKey, this.User1Information.CalendarCollectionId, addCalendar);
            SyncResponse syncAddCalendarResponse = this.CMDAdapter.Sync(syncAddCalendarRequest);

            // Get data from response
            Response.SyncCollections syncCollections = (Response.SyncCollections)syncAddCalendarResponse.ResponseData.Item;
            Response.SyncCollectionsCollectionResponses syncResponses = null;
            for (int index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
            {
                if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Responses)
                {
                    syncResponses = (Response.SyncCollectionsCollectionResponses)syncCollections.Collection[0].Items[index];
                    break;
                }
            }

            Site.Assert.AreEqual(1, syncResponses.Add.Length, "User only upload one calendar item");
            int statusCode = int.Parse(syncResponses.Add[0].Status);
            Site.Assert.AreEqual(1, statusCode, "If upload calendar item successful, server should return status 1");
        }
        #endregion
    }
}