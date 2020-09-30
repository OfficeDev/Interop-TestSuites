namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to test the MeetingResponse command.
    /// </summary>
    [TestClass]
    public class S06_MeetingResponseCommand : TestSuiteBase
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

        #region MSASAIRS_S06_TC01_MeetingResponseWithPlainTextBody
        /// <summary>
        /// This case is designed to test the MeetingResponse command. 
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S06_TC01_MeetingResponseWithPlainTextBody()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

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
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
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
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);

            Request.MeetingResponseRequestSendResponse sendResoponse=new Request.MeetingResponseRequestSendResponse()
            {
                Body=new Request.Body()
                {
                    Type= 1,
                    Data="Accept this meeting."
                }
            };

            // Accept the meeting request
            // Create a meeting response request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = 1,
                CollectionId = this.User2Information.InboxCollectionId,
                RequestId = meetingRequestEmail.ServerId,
                SendResponse = sendResoponse,
            };

            // Create a meeting response request
            MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse response = this.ASAIRSAdapter.MeetingResponse(meetingRequest);

            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                1331,
                @"[In Body] When the Body element is a child of the meetingresponse:SendResponse element [or the composemail:SmartForward element], it has only the child elements Type and Data.");

            this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                1333,
                @"[In Body] The Body element is a child of the meetingresponse:SendResponse element and the composemail:SmartForward element only when protocol version 16.0 is used.");

            // Because the Type element is 1 and client call the MeetingResponse command successful.
            // So R1400 and R14000918 will be verified.
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
            {
                this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                1400,
                @"[In Type (Body)] For calendar items in protocol version 16.0, the only valid values for this element [Type] is 1 (plain text).");
            }
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                14000918,
                @"[In Type (Body)] For calendar items in protocol version 16.1, the only valid values for this element [Type] is 1 (plain text).");
            }
            #endregion

            #region Call Sync command to get the calendar item.
            SyncStore getCalendarItemsResult = this.GetSyncResult(subject, this.User2Information.CalendarCollectionId, null);
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync calendarResult = TestSuiteHelper.GetSyncAddItem(getCalendarItemsResult, subject);
            Site.Assert.IsNotNull(calendarResult.Calendar.BusyStatus, "Element BusyStatus should be present.");
            #endregion
        }
        #endregion

        #region MSASAIRS_S06_TC02_MeetingResponseWithHTMLBody
        /// <summary>
        /// This case is designed to test the MeetingResponse command. 
        /// </summary>
        [TestCategory("MSASAIRS"), TestMethod()]
        public void MSASAIRS_S06_TC02_MeetingResponseWithHTMLBody()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The Body element under meetingresponse:SendResponse element is supported when the MS-ASProtocolVersion header is set to 16.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

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
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync resultItem = TestSuiteHelper.GetSyncAddItem(getChangeResult, subject);
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
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync meetingRequestEmail = TestSuiteHelper.GetSyncAddItem(syncChangeResult, subject);

            Request.MeetingResponseRequestSendResponse sendResoponse = new Request.MeetingResponseRequestSendResponse()
            {
                Body = new Request.Body()
                {
                    Type = 2,
                    Data = "<html><head></head><body>Accept this meeting.</body></html>"
                }
            };

            // Accept the meeting request
            // Create a meeting response request item
            Request.MeetingResponseRequest meetingResponseRequestItem = new Request.MeetingResponseRequest
            {
                UserResponse = 1,
                CollectionId = this.User2Information.InboxCollectionId,
                RequestId = meetingRequestEmail.ServerId,
                SendResponse = sendResoponse,
            };

            // Create a meeting response request
            MeetingResponseRequest meetingRequest = Common.CreateMeetingResponseRequest(new Request.MeetingResponseRequest[] { meetingResponseRequestItem });
            MeetingResponseResponse response = this.ASAIRSAdapter.MeetingResponse(meetingRequest);

            // Because the Type element is 2 and client call the MeetingResponse command successful.
            // So R1401 and R14010918 will be verified.
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0"))
            {
                this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                1401,
                @"[In Type (Body)] For calendar items in protocol version 16.0, the only valid values for this element [Type] is 2 (HTML).");
            }
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))
            {
                this.Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(response.ResponseData.Result[0].Status),
                14010918,
                @"[In Type (Body)] For calendar items in protocol version 16.1, the only valid values for this element [Type] is 2 (HTML).");
            }
            #endregion

            #region Call Sync command to get the calendar item.
            SyncStore getCalendarItemsResult = this.GetSyncResult(subject, this.User2Information.CalendarCollectionId, null);
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync calendarResult = TestSuiteHelper.GetSyncAddItem(getCalendarItemsResult, subject);
            Site.Assert.IsNotNull(calendarResult.Calendar.BusyStatus, "Element BusyStatus should be present.");
            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Add a meeting to the server.
        /// </summary>
        /// <param name="calendarCollectionId">The collectionId of the folder which the item should be added.</param>
        /// <param name="elementsToValueMap">The key and value pairs of common meeting properties.</param>
        private void SyncAddMeeting(string calendarCollectionId, Dictionary<Request.ItemsChoiceType8, object> elementsToValueMap)
        {
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

            string iniSyncKey = this.GetInitialSyncKey(calendarCollectionId);
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncAddRequest(iniSyncKey, calendarCollectionId, applicationData);

            SyncStore syncAddResponse = this.ASAIRSAdapter.Sync(syncAddRequest);
            Site.Assert.AreEqual<int>(
                1,
                int.Parse(syncAddResponse.AddResponses[0].Status),
                "The sync add operation should be successful.");
        }

        /// <summary>
        /// Send a meeting request email.
        /// </summary>
        /// <param name="subject">The subject of email</param>
        /// <param name="calendar">The meeting calendar</param>
        private void SendMeetingRequest(string subject, Calendar calendar)
        {
            string emailBody = Common.GenerateResourceName(Site, "content");
            string icalendarFormatContent = TestSuiteHelper.CreateiCalendarFormatContent(calendar);

            string meetingEmailMime = TestSuiteHelper.CreateMeetingRequestMime(
                calendar.OrganizerEmail,
                calendar.Attendees.Attendee[0].Email,
                subject,
                emailBody,
                icalendarFormatContent);

            Request.SendMail sendMail = new Request.SendMail();

            sendMail.ClientId = Guid.NewGuid().ToString("N");
            sendMail.Mime = meetingEmailMime;

            SendMailRequest sendMailRequest = Common.CreateSendMailRequest();
            sendMailRequest.RequestData = sendMail;

            this.SwitchUser(this.User1Information, false);
            SendMailResponse response = this.ASAIRSAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command success.");
        }

        /// <summary>
        /// Get the specified email item.
        /// </summary>
        /// <param name="emailSubject">The subject of the email item.</param>
        /// <param name="folderCollectionId">The serverId of the default folder.</param>
        /// <param name="bodyPreference">The preference information related to the type and size of information that is returned from fetching.</param>
        /// <returns>The result of getting the specified email item.</returns>
        private SyncStore GetSyncResult(string emailSubject, string folderCollectionId, Request.BodyPreference bodyPreference)
        {
            SyncStore syncItemResult;
            Microsoft.Protocols.TestSuites.Common.DataStructures.Sync item = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);

                // Get the new added email item
                syncItemResult = this.SyncChanges(this.GetInitialSyncKey(folderCollectionId), folderCollectionId, bodyPreference);
                if (syncItemResult != null && syncItemResult.CollectionStatus == 1)
                {
                    item = TestSuiteHelper.GetSyncAddItem(syncItemResult, emailSubject);
                }

                counter++;
            }
            while ((syncItemResult == null || item == null) && counter < retryCount);

            Site.Assert.IsNotNull(item, "The email item with subject {0} should be found. Retry count: {1}", emailSubject, counter);

            // Verify sync result
            Site.Assert.AreEqual<byte>(
                1,
                syncItemResult.CollectionStatus,
                "If the Sync command executes successfully, the status in response should be 1.");

            return syncItemResult;
        }

        /// <summary>
        /// Sync changes between client and server
        /// </summary>
        /// <param name="syncKey">The synchronization key returned by last request.</param>
        /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body</param>
        /// <returns>Return change result</returns>
        private SyncStore SyncChanges(string syncKey, string collectionId, Request.BodyPreference bodyPreference)
        {
            // Get changes from server use initial syncKey
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, bodyPreference);
            SyncStore syncResult = this.ASAIRSAdapter.Sync(syncRequest);

            return syncResult;
        }
        #endregion
    }
}
