namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is used to test the SmartReply command.
    /// </summary>
    [TestClass]
    public class S18_SmartReply : TestSuiteBase
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
        /// This test case is used to verify the replied mail contains original message.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S18_TC01_SmartReply_ContainOriginalMessage()
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

            #region Call Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncChangeResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

            // Get User2 original mail content
            string originalServerId = TestSuiteBase.FindServerId(syncChangeResponse, "Subject", meetingRequestSubject);
            string originalContent = TestSuiteBase.GetDataFromResponseBodyElement(syncChangeResponse, originalServerId);
            #endregion

            #region Record user name, folder collectionId and item subject that are useed in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Call SmartReply command to reply to messages without instanceID in request
            string smartReplySubject = string.Format("REPLY: {0}", meetingRequestSubject);
            SmartReplyRequest smartReplyRequest = this.CreateDefaultReplyRequest(smartReplySubject, calendarItemID);

            // Add elements to smartReplyRequest
            smartReplyRequest.RequestData.Source.InstanceId = null;
            smartReplyRequest.RequestData.Source.FolderId = this.User2Information.CalendarCollectionId;
            smartReplyRequest.RequestData.Source.ItemId = calendarItemID;
            this.CMDAdapter.SmartReply(smartReplyRequest);
            #endregion

            #region Call Sync command to sync user1 mailbox changes
            this.SwitchUser(this.User1Information);
            SyncResponse syncResponseOnUserOne = this.GetMailItem(this.User1Information.InboxCollectionId, smartReplySubject);

            // Get replied mail content
            string replyMailServerId = TestSuiteBase.FindServerId(syncResponseOnUserOne, "Subject", smartReplySubject);
            string replyMailContent = TestSuiteBase.GetDataFromResponseBodyElement(syncResponseOnUserOne, replyMailServerId);
            #endregion

            #region Record user name, folder collectionId and item subject that are useed in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, smartReplySubject);
            #endregion

            #region Verify Requirements MS-ASCMD_R5420, MS-ASCMD_R580, MS-ASCMD_R569

            if (Common.IsRequirementEnabled(5420, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5420");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5420
                // If SmartReply command executes successful, the reply mail content contains the original mail content
                Site.CaptureRequirementIfIsTrue(
                     replyMailContent.Contains(originalContent),
                     5420,
                     @"[In Appendix A: Product Behavior] If SmartReply is applied to a recurring meeting and the InstanceId element is absent, the implementation does reply for the entire recurring meeting. (Exchange 2007 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R580");

            // Verify MS-ASCMD requirement: MS-ASCMD_R580
            Site.CaptureRequirementIfIsTrue(
                 replyMailContent.Contains(originalContent),
                 580,
                 @"[In SmartReply] The full text of the original message is retrieved and sent by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R569");

            // Verify MS-ASCMD requirement: MS-ASCMD_R569
            Site.CaptureRequirementIfIsTrue(
                 replyMailContent.Contains(originalContent),
                 569,
                 @"[In SmartReply] The SmartReply command is used by clients to reply to messages without retrieving the full, original message from the server.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the server returns an empty response, when mail is replied successfully.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S18_TC02_SmartReply_Success()
        {
            #region Call SendMail command to send one plain text email to user2.
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            this.SendPlainTextEmail(null, emailSubject, this.User1Information.UserName, this.User2Information.UserName, null);
            #endregion

            #region Call Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncChangeResponse = this.GetMailItem(this.User2Information.InboxCollectionId, emailSubject);
            string originalServerId = TestSuiteBase.FindServerId(syncChangeResponse, "Subject", emailSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are useed in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Call SmartReply command to reply to messages without retrieving the full, original message from the server
            string smartReplySubject = string.Format("REPLY: {0}", emailSubject);
            SmartReplyRequest smartReplyRequest = this.CreateDefaultReplyRequest(smartReplySubject, originalServerId);

            SmartReplyResponse smartReplyResponse = this.CMDAdapter.SmartReply(smartReplyRequest);

            #endregion

            #region Call Sync command to sync user1 mailbox changes
            this.SwitchUser(this.User1Information);
            this.GetMailItem(this.User1Information.InboxCollectionId, smartReplySubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are useed in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, smartReplySubject);
            #endregion

            #region Verify Requirements MS-ASCMD_R605, MS-ASCMD_R5776
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R605");

            // Verify MS-ASCMD requirement: MS-ASCMD_R605
            // If the message was sent successfully, the server won't return any XML data, then MS-ASCMD_R605 is verified.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                smartReplyResponse.ResponseDataXML,
                605,
                @"[In SmartReply] If the message was sent successfully, the server returns an empty response.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5776");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5776
            // If the message was sent successfully, the server won't return any XML data, then MS-ASCMD_R776 is verified.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                smartReplyResponse.ResponseDataXML,
                5776,
                @"[In Status(SmartForward and SmartReply)] If the [SmartForward command request or] SmartReply command request succeeds, no XML body is returned in the response.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the value of the InstanceId element is invalid, the server responds with Status value 104.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S18_TC03_SmartReply_Status104()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes.
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 creates SmartReply request with invalid InstanceId value, then calls SmartReply command.
            // Set instanceID with format that is not the same as required "2010-03-20T22:40:00.000Z".
            string instanceID = DateTime.Now.ToString();
            string smartReplySubject = string.Format("REPLY: {0}", meetingRequestSubject);
            SmartReplyRequest smartReplyRequest = this.CreateDefaultReplyRequest(smartReplySubject, calendarItemID);

            // Add instanceId element to smartReplyRequest.
            smartReplyRequest.RequestData.Source.InstanceId = instanceID;
            smartReplyRequest.RequestData.Source.FolderId = this.User2Information.CalendarCollectionId;
            SmartReplyResponse smardReplyResponse = this.CMDAdapter.SmartReply(smartReplyRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R5777, MS-ASCMD_R578
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5777");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5777
            // If server returns status code which has value, that indicates the type of failure then MS-ASCMD_R5777 is verified.
            Site.CaptureRequirementIfIsNotNull(
                smardReplyResponse.ResponseData.Status,
                5777,
                @"[In Status(SmartForward and SmartReply)] If the [SmartForward command request or ] SmartReply command request fails, the Status element contains a code that indicates the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R578");

            // Verify MS-ASCMD requirement: MS-ASCMD_R578
            Site.CaptureRequirementIfAreEqual<string>(
                "104",
                smardReplyResponse.ResponseData.Status,
                578,
                @"[In SmartReply] If the value of the InstanceId element is invalid, the server responds with Status element (section 2.2.3.162.15) value 104, as specified in section 2.2.4.");
            #endregion
        }
        #endregion

        #region Private method
        /// <summary>
        /// Create default SmartReply request.
        /// </summary>
        /// <param name="replySubject">Reply email subject.</param>
        /// <param name="originalServerId">The  server ID of the original email.</param>
        /// <returns>Smart Reply request.</returns>
        private SmartReplyRequest CreateDefaultReplyRequest(string replySubject, string originalServerId)
        {
            string smartReplyFromUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string smartReplyToUser = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string smartReplyContent = Common.GenerateResourceName(Site, "reply:body");
            SmartReplyRequest smartReplyRequest = TestSuiteBase.CreateSmartReplyRequest(this.User2Information.InboxCollectionId, originalServerId, smartReplyFromUser, smartReplyToUser, string.Empty, string.Empty, replySubject, smartReplyContent);
            return smartReplyRequest;
        }
        #endregion
    }
}