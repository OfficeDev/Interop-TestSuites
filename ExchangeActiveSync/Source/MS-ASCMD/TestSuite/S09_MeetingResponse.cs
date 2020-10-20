namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is used to test the MeetingResponse command.
    /// </summary>
    [TestClass]
    public class S09_MeetingResponse : TestSuiteBase
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

        #region Test Cases
        /// <summary>
        /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC01_MeetingResponse_AcceptMeeting()
        {
            #region User1 calls SendMail command to send one meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);

            // Sync Inbox folder
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string serverIDForMeetingRequest = TestSuiteBase.FindServerId(syncResponse, "Subject", meetingRequestSubject);
            Response.MeetingRequest meetingRequest = (Response.MeetingRequest)TestSuiteBase.GetElementValueFromSyncResponse(syncResponse, serverIDForMeetingRequest, Response.ItemsChoiceType8.MeetingRequest);

            // Sync Calendar folder
            SyncResponse syncCalendarBeforeMeetingResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemServerID = TestSuiteBase.FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
            string messageClass = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncResponse, serverIDForMeetingRequest, Response.ItemsChoiceType8.MessageClass);
            #endregion

            #region Verify Requirements MS-ASCMD_R5068, MS-ASCMD_R5085, MS-ASCMD_R5828
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5068");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5068
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Schedule.Meeting.Request",
                messageClass,
                5068,
                @"[In Receiving and Accepting Meeting Requests] The message contains an email:MessageClass element (as specified in [MS-ASEMAIL] section 2.2.2.41) that has a value of ""IPM.Schedule.Meeting.Request"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5085");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5085
            // User calls Sync command to get Inbox folder changes, if Sync response contains MeetingRequest element then MS-ASCMD_R5085 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingRequest,
                5085,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 2*:] The server responds with airsync:Add elements (section 2.2.3.7.2) for items in the Inbox collection, including a meeting request item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5828");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5828
            // User calls Sync command to get Inbox folder changes, if Sync response contains MeetingRequest element then MS-ASCMD_R5828 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingRequest,
                5828,
                @"[In Receiving and Accepting Meeting Requests] Its [The message's] airsync:ApplicationData element (section 2.2.3.11) contains an email:MeetingRequest element (as specified in [MS-ASEMAIL] section 2.2.2.40).");
            #endregion

            #region Call method MeetingResponse to accept the meeting request in the user2's Inbox folder.
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

            // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R5082, MS-ASCMD_R4180, MS-ASCMD_R830, MS-ASCMD_R831, MS-ASCMD_R3843, MS-ASCMD_R5071, MS-ASCMD_R5089, MS-ASCMD_R5723

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5082");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5082
            // User calls Sync command with the SyncKey element value of 0 for Inbox folder, if this.LastSynKey is not null, means Sync operation success and server returned SyncKey value, then MS-ASCMD_R830 is verified.
            Site.CaptureRequirementIfIsNotNull(
                this.LastSyncKey,
                5082,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 1:] The server responds with the airsync:SyncKey for the collection, to be used in successive synchronizations.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4180");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4180
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                meetingResponseResponse.ResponseData.Result[0].Status,
                4180,
                @"[In Status(MeetingResponse)] [When the scope is Global], [the cause of the status value 1 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R830");

            // Verify MS-ASCMD requirement: MS-ASCMD_R830
            // If user accept the meeting request, server will return CalendarId element in response.
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].CalendarId,
                830,
                @"[In CalendarId] The CalendarId element is included in the MeetingResponse command response that is sent to the client if the meeting request was not declined.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R831");

            // Verify MS-ASCMD requirement: MS-ASCMD_R831
            // If MeetingResponse command executes successfully, the server will return calendarId element in the response, then MS-ASCMD_R831 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].CalendarId,
                831,
                @"[In CalendarId] If the meeting is accepted [or tentatively accepted], the server adds a new item to the calendar and returns its server ID in the CalendarId element in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3843");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3843
            // If the meeting request is accepted, the server will respond calendarId element in response, which is the calendar item server ID, then MS-ASCMD_R4383 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].CalendarId,
                3843,
                @"[In Result(MeetingResponse)] If the meeting request is accepted, the server ID of the calendar item is also returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5071");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5071
            // If the meeting request is accepted, the server will response calendarId element in response, which is the calendar item server ID, then MS-ASCMD_R5071 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].CalendarId,
                5071,
                @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is accepted [or is tentatively accepted], the server will add or update the corresponding calendar item and return its server ID in the meetingresponse:CalendarId element (section 2.2.3.18) of the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5089");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5089
            Site.CaptureRequirementIfIsTrue(
                meetingResponseResponse.ResponseData.Result[0].CalendarId != null && int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) != 0,
                5089,
                @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 4:] The server sends a response that contains the MeetingResponse command request status along with the ID of the calendar item that corresponds to this meeting request if the meeting was not declined.");

            if (Common.IsRequirementEnabled(5723, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5723");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5723
                Site.CaptureRequirementIfIsTrue(
                    meetingResponseResponse.ResponseData.Result[0].CalendarId != null && int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) != 0,
                    5723,
                    @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to accept a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Sync Calendar folder change and get accepted calendar item's serverID
            SyncResponse syncCalendarResponseAfterAcceptMeeting = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponseAfterAcceptMeeting, "Subject", meetingRequestSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);
            #endregion

            // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
            {
                // Get calendar item responseType value
                uint responseTypeBeforeMeetingResponse = (uint)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarBeforeMeetingResponse, calendarItemServerID, Response.ItemsChoiceType8.ResponseType);
                uint responseTypeAfterMeetingResponse = (uint)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponseAfterAcceptMeeting, calendarItemID, Response.ItemsChoiceType8.ResponseType);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5094");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5094
                Site.CaptureRequirementIfIsTrue(
                    responseTypeBeforeMeetingResponse != responseTypeAfterMeetingResponse && meetingResponseResponse.ResponseData.Result[0].CalendarId != null,
                    5094,
                    @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 6:] The server responds with any changes to the Calendar folder caused by the last synchronization and the new calendar item for the accepted meeting.");
            }

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "12.1")
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5094");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5094
                // If the meeting was accepted, server should return calendarId in response
                Site.CaptureRequirementIfIsTrue(
                    meetingResponseResponse.ResponseData.Result[0].CalendarId != null,
                    5094,
                    @"[In Receiving and Accepting Meeting Requests] [Command sequence for receiving and accepting meeting requests., order 6:] The server responds with any changes to the Calendar folder caused by the last synchronization and the new calendar item for the accepted meeting.");
            }
        }

        /// <summary>
        /// This test case is used to verify the response that has no CalendarId, when meeting response is declined.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC02_MeetingResponse_DeclineMeeting()
        {
            #region User1 calls SendMail command to send one meeting request to user2

            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            string serverIDForMeetingRequest = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to decline the meeting request in the Inbox folder.
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(3, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

            // If the user declines the meeting request, the meeting request mail will be deleted and no calendar item will be created.
            MeetingResponseResponse responseMeetingResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            string itemServerIDInDeletefolder = this.GetItemServerIdFromSpecialFolder(this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);
            Site.Assert.IsNotNull(itemServerIDInDeletefolder, "If user decline the meeting request, the meeting request mail should be deleted");
            this.DeleteAll(this.User2Information.DeletedItemsCollectionId);

            SyncResponse syncInboxFolder = this.SyncChanges(this.User2Information.InboxCollectionId);
            string itemServerIDInInboxFolder = TestSuiteBase.FindServerId(syncInboxFolder, "Subject", meetingRequestSubject);

            if (Common.IsRequirementEnabled(5725, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5725");

                // If call the MeetingResponse command to decline a meeting request, the original meeting request item will move from Inbox folder to DeleteItems folder.
                // Verify MS-ASCMD requirement: MS-ASCMD_R5725
                Site.CaptureRequirementIfIsTrue(
                    itemServerIDInDeletefolder != null && itemServerIDInInboxFolder == null,
                    5725,
                    @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to decline a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Verify Requirements MS-ASCMD_R837, MS-ASCMD_R5072
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R837");

            // Verify MS-ASCMD requirement: MS-ASCMD_R837
            // If user declined the meeting the server response will not return calendarId element
            Site.CaptureRequirementIfIsNull(
                responseMeetingResponse.ResponseData.Result[0].CalendarId,
                837,
                @"[In CalendarId] If the meeting is declined, the response does not contain a CalendarId element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5072");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5072
            // If user declined the meeting the server response will not return calendarId element
            Site.CaptureRequirementIfIsNull(
                responseMeetingResponse.ResponseData.Result[0].CalendarId,
                5072,
                @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is declined, the response will not contain a meetingresponse:CalendarId element because the server will delete the corresponding calendar item.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify returned status value is 2, when userResponse is invalid.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC03_MeetingResponse_InvalidMeeting()
        {
            #region User1 calls SendMail command to send one meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email
            this.SwitchUser(this.User2Information);
            string serverIDForMeetingRequest = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, meetingRequestSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse with invalid UserResponse element.
            // Set invalid UserResponse value "5"
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(5, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);
            MeetingResponseResponse responseMeetingResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4182");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4182
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(responseMeetingResponse.ResponseData.Result[0].Status),
                4182,
                @"[In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The client has sent a malformed or invalid item.");
        }

        /// <summary>
        /// This test case is used to verify RequestId is present in MeetingResponse command response if it was present in the corresponding MeetingResponse command request.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC04_MeetingResponse_RequestID()
        {
            #region User1 calls SendMail command to send one meeting request to user2

            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email.
            this.SwitchUser(this.User2Information);
            string requestID = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to tentatively accept the meeting request in Inbox folder.
            // Set UserResponse value 2 to tentatively accepted
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(2, this.User2Information.InboxCollectionId, requestID, string.Empty);
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
            string itemServerID = this.GetItemServerIdFromSpecialFolder(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);

            // If user tentatively accepted the meeting, the calendar item will be found in Calendar folder.
            if (Common.IsRequirementEnabled(5724, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5724");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5724
                Site.CaptureRequirementIfIsTrue(
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status) == 1 && itemServerID != null,
                    5724,
                    @"[In Appendix A: Product Behavior] Implementation does use the MeetingResponse command to tentatively accept a meeting request in the user's Inbox folder. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Verify Requirements MS-ASCMD_R5374, MS-ASCMD_R3807
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5374");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5374
            Site.CaptureRequirementIfAreEqual(
                requestID,
                meetingResponseResponse.ResponseData.Result[0].RequestId,
                5374,
                @"[In RequestId] The RequestId element is present in MeetingResponse command responses only if it was present in the corresponding MeetingResponse command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3807");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3807
            // If server response contains RequestId element, then MS-ASCMD_3807 is verified.
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].RequestId,
                3807,
                @"[In RequestId] The RequestId element is also returned in the response to the client along with the status of the user's response to the meeting request.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MeetingResponse command response has status equals 2, when the request is referencing an item other than a meeting request, e-mail or calendar item.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC05_MeetingResponse_ResponseNonMeeting()
        {
            #region User1 calls SendMail command to send one meeting request to user2

            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region user2 get new added meeting request email
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to tentatively accept the meeting request with invalid request.
            // Create request with invalid RequestID value
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(2, this.User2Information.InboxCollectionId, "InvalidValue", string.Empty);
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4183");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4183
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                4183,
                @"[In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The request is referencing an item other than a meeting request, email, or calendar item.");
        }

        /// <summary>
        /// This test case is used to verify if the InstanceId is not a specified meeting request, server should return the status value is 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC06_MeetingResponse_NonExistInstanceId()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region user2 get new added meeting request email
            this.SwitchUser(this.User2Information);
            string serverIDForMeetingResponse = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, meetingRequestSubject);
            this.GetItemServerIdFromSpecialFolder(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to decline the meeting request in Inbox folder with invalid request.
            // Create invalid MeetingResponse request with invalid InstanceID value
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(3, this.User2Information.InboxCollectionId, serverIDForMeetingResponse, "InvalidValue");
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4186");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4186
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                4186,
                @"[In Status(MeetingResponse)] [In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The InstanceId element specifies a nonexistent instance or is null.");
        }

        /// <summary>
        /// This test case is used to verify if there are more than 100 Request elements listed in the MeetingResponse command request, the server will return status 103.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC07_MeetingResponse_Status103()
        {
            #region User1 calls SendMail command to send one meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region User2 calls Sync command to get new added meeting request email
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            string meetingRequestServerID = this.GetItemServerIdFromSpecialFolder(this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 creates MeetingResponse command request with 101 Request elements

            int totalCount = 101;
            List<string> serverIDList = new List<string>();
            for (int requestIndex = 0; requestIndex < totalCount; requestIndex++)
            {
                serverIDList.Add(meetingRequestServerID);
            }

            MeetingResponseRequest meetingResponseRequest = CreateMultiMeetingResponseRequest(1, this.User2Information.InboxCollectionId, serverIDList, string.Empty);

            // Send MeetingResponse command request with 101 Request elements
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            if (Common.IsRequirementEnabled(5672, this.Site) || Common.IsRequirementEnabled(5670, this.Site))
            {
                TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
                TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5670");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5670
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                    5670,
                    @"[In Appendix A: Product Behavior] Implementation does not limit the number of elements in command requests and not return the specified error if the limit is exceeded. (<118> Section 3.1.5.8: Exchange 2007 SP1 and Exchange 2010 do not limit the number of elements in command requests.) ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5672");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5672
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                    5672,
                    @"[In Appendix A: Product Behavior] Implementation does not limit the number of elements in command requests. (<119> Section 3.1.5.8: Exchange 2007 SP1 and Exchange 2010 do not limit the number of elements in command requests. )");
            }

            if (Common.IsRequirementEnabled(5671, this.Site) || Common.IsRequirementEnabled(5673, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5671");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5671
                Site.CaptureRequirementIfAreEqual<int>(
                    103,
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                    5671,
                    @"[In Appendix A: Product Behavior] Implementation does limit the number of elements in command requests and return the specified error if the limit is exceeded. (<118> Section 3.1.5.8: Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 do limit the number of elements in command requests.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5673");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5673
                Site.CaptureRequirementIfAreEqual<int>(
                    103,
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                    5673,
                    @"[In Appendix A: Product Behavior] Update Rollup 6 for implementation does use the specified limit values by default but can be configured to use different values. (<119> Section 3.1.5.8: Update Rollup 6 for Exchange 2010 SP2 and Exchange 2013 use the specified limit values by default but can be configured to use different values.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5650");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5650
                Site.CaptureRequirementIfAreEqual<int>(
                    103,
                    int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                    5650,
                    @"[In Limiting Size of Command Requests] In MeetingResponse (section 2.2.2.9) command request, when the limit value of Request element is bigger than 100 (minimum 1, maximum 2,147,483,647), the error returned by server is Status element (section 2.2.3.162.8) value of 103.");
            }
        }

        /// <summary>
        /// This test case is used to verify if the InstanceId element specifies an email meeting request item, the server returns status code 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC08_MeetingResponse_RecurringMeetingInstanceIDInvalid()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);

            // Get the meeting request mail from Inbox folder.
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);

            // Get the calendar item from Calendar folder
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            string startTime = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to accept the meetingRequest with Instance element referring to email meeting request item in MeetingResponseRequest
            DateTime calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();

            // Set invalid instanceID randomly.            
            string instanceID = new Random().ToString();
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.InboxCollectionId, calendarItemID, instanceID);
            
            // Send MeetingResponseRequest with instanceID specifies a email meeting
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4185");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4185
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                4185,
                @"[In Status(MeetingResponse)] [In Status(MeetingResponse)] [When the scope is Item], [the cause of the status value 2 is] The InstanceId element (section 2.2.3.78.1) specifies an email meeting request item.");
        }

        /// <summary>
        /// This test case is used to verify if the InstanceId element value specifies a non-recurring meeting, the server responds with a Status element value of 146.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC09_MeetingResponse_Status146()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send one single meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region User2 calls Sync command to get new added meeting request
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);

            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);

            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to accept the meetingRequest with InstanceId element
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            string startTime = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
            DateTime calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();
            string instanceID = calendarStartTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");

            // Send MeetingResponseRequest with instanceID specifies a non-recurring meeting
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.CalendarCollectionId, calendarItemID, instanceID);
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3196");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3196
            Site.CaptureRequirementIfAreEqual<int>(
                146,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                3196,
                @"[In InstanceId(MeetingResponse)] If the InstanceId element value specifies a non-recurring meeting, the server responds with a Status element value of 146.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4919");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4919
            Site.CaptureRequirementIfAreEqual<int>(
                146,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                4919,
                @"[In Common Status Codes] [The meaning of the status value 146 is] The request tried to forward an occurrence of a meeting that has no recurrence.");
        }

        /// <summary>
        /// This test is used to verify implementation does use the MeetingResponse command to tentatively accept a meeting request in the user's Inbox folder
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC10_MeetingResponse_TentativeAcceptMeeting()
        {
            #region User1 calls SendMail command to send one meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncInboxResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string inboxItemID = TestSuiteBase.FindServerId(syncInboxResponse, "Subject", meetingRequestSubject);

            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

            // Get calendar item responseType value before meetingResponse
            // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
            uint responseTypeBeforeMeetingResponse = 0;
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
            {
                responseTypeBeforeMeetingResponse = (uint)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);
            }

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to tentative accept the meeting
            // Set to tentatively accept one of recurring meeting request instance
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(2, this.User2Information.InboxCollectionId, inboxItemID, null);
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);

            SyncResponse syncCalendarResponseAfterMeetingResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region Verify Requirements MS-ASCMD_R5790, MS-ASCMD_R5678
            // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
            {
                // Get calendar item responseType value
                uint responseTypeAfterMeetingResponse = (uint)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponseAfterMeetingResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5790");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5790
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    responseTypeBeforeMeetingResponse,
                    responseTypeAfterMeetingResponse,
                    5790,
                    @"[In Receiving and Accepting Meeting Requests] If the response to the meeting is [accepted or is] tentatively accepted, the server will add or update the corresponding calendar item and return its server ID in the meetingresponse:CalendarId element (section 2.2.3.18) of the response.");
            }

            // If meetingResponse response returns new calendarId element, then MS-ASCMD_R5678 is verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5678");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5678
            Site.CaptureRequirementIfIsNotNull(
                meetingResponseResponse.ResponseData.Result[0].CalendarId,
                5678,
                @"[In CalendarId] If the meeting is [accepted or] tentatively accepted, the server adds a new item to the calendar and returns its server ID in the CalendarId element in the response.");
            #endregion
        }

        /// <summary>
        /// This test is used to verify implementation does use the MeetingResponse command to accept a meeting request in the user's Calendar folder.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC11_MeetingResponse_AcceptMeetingInCalendar()
        {
            #region User1 calls SendMail command to send one meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress,null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region User2 calls Sync and FolderSync commands to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string calendarItemID = this.GetItemServerIdFromSpecialFolder(this.User2Information.CalendarCollectionId, meetingRequestSubject);

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to accept the meetingRequest
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.CalendarCollectionId, calendarItemID, null);
            this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if the InstanceId element value specified is not in the proper format, the server responds with a Status value of 104.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC12_MeetingResponse_Status104()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The InstanceId element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            #region User1 calls SendMail command to send one recurring meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.SendWeeklyRecurrenceMeetingRequest(meetingRequestSubject, attendeeEmailAddress);
            #endregion

            #region User2 calls Sync command to get new added meeting request
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);
            this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);

            // Record related items that need to be cleaned up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to accept the meetingRequest with wrong format Instance element in MeetingResponseRequest
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);
            string startTime = (string)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.StartTime);
            DateTime calendarStartTime = Common.GetNoSeparatorDateTime(startTime).ToUniversalTime();

            // Set instanceID using month-day-yearThour:minute:second.milsecondZ format which is different from required format "2010-03-20T22:40:00.000Z".
            string instanceID = calendarStartTime.ToString("MM-dd-yyyyTHH:mm:ss.fffZ");

            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.CalendarCollectionId, calendarItemID, instanceID);

            // Send MeetingResponseRequest with instanceID specifies a non-recurring meeting
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            #endregion

            #region Verify Requirements MS-ASCMD_R3195, MS-ASCMD_R4819
            // If the InstanceId element value is not in the proper format, server returns status code 104 which means the value is in invalid format, then MS-ASCMD_R3195, MS-ASCMD_R4819 are verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3195");

            // Verify MS-ASCMD requirement: MS-ASCMD_R3195
            Site.CaptureRequirementIfAreEqual<int>(
                104,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                3195,
                @"[In InstanceId(MeetingResponse)] If the InstanceId element value specified is not in the proper format, the server responds with a Status element (section 2.2.3.162.8) value of 104.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4819");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4819
            Site.CaptureRequirementIfAreEqual<int>(
                104,
                int.Parse(meetingResponseResponse.ResponseData.Result[0].Status),
                4819,
                @"[In Common Status Codes] [The meaning of the status value 104 is] The request contains a timestamp that could not be parsed into a valid date and time.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that server sends a substitute meeting invitation email.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC13_MeetingResponse_SubstituteMeetingInvitationEmail()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5822, this.Site), "Exchange Server 2013 and above support Substitute Meeting Invitation.");

            #region User1 calls SendMail command to send one meeting request to user7
            string originalSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User7Information.UserName, this.User7Information.UserDomain);
            Calendar calendar = this.CreateCalendar(originalSubject, attendeeEmailAddress,null);

            // Send a meeting request email to user7
            this.SendMeetingRequest(originalSubject, calendar);
            #endregion

            #region Get new added meeting request email in user7 mailbox
            // Switch to user7 mailbox
            this.SwitchUser(this.User7Information);

            // Sync Inbox folder
            this.GetMailItem(this.User7Information.InboxCollectionId, originalSubject);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R7478");

            // If the protocol version is 16.0, nothing will be done in the method SendMeetingRequest.
            // So when the assertion in GetMailItem succeeds, R7478 can be verified.
            Site.CaptureRequirement(
                7478,
                @"[In Creating a Meeting or Appointment] In protocol version 16.0, the server will send meeting requests to the attendees automatically while processing the Sync command request that creates the meeting.");

            // Sync Calendar folder
            this.GetMailItem(this.User7Information.CalendarCollectionId, originalSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User7Information, this.User7Information.CalendarCollectionId, originalSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User7Information, this.User7Information.InboxCollectionId, originalSubject);
            #endregion

            #region Get substitute invitation email in user8 mailbox
            // Switch to user8 mailbox, user8 is delegate of user7
            this.SwitchUser(this.User8Information);

            // Sync Inbox folder
            SyncResponse substituteSyncResponse = this.GetSubstituteMailItem(this.User8Information.InboxCollectionId, originalSubject);
            string substituteInvitationEmailServerId = TestSuiteBase.FindServerId(substituteSyncResponse, "ThreadTopic", originalSubject);
            #endregion

            #region Verify Requirements related receiving and accepting a meeting request
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5613");

            // If the element Body is not null, the response contains the informations of meeting, this requirement can be verified.
            // Verify MS-ASCMD requirement: MS-ASCMD_R5613
            Site.CaptureRequirementIfIsNotNull(
                TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Body),
                5613,
                @"[In Substitute Meeting Invitation Email] The value of element airsyncbase:Body is summary of meeting details.");

            string messageClass = (string)TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.MessageClass);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5822");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5822
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Note",
                messageClass,
                5822,
                @"[In Appendix A: Product Behavior] Implementation does return substitute meeting invitation email messages. (Exchange 2013 and above follow this behavior.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5614");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5614
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Note",
                messageClass,
                5614,
                @"[In Substitute Meeting Invitation Email] The value of element email:MessageClass is set to ""IPM.Note"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5604");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5604
            Site.CaptureRequirementIfIsTrue(
                ((string)TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.To)).ToLower(System.Globalization.CultureInfo.InvariantCulture).Contains(this.User8Information.UserName.ToLower(System.Globalization.CultureInfo.InvariantCulture)),
                5604,
                @"[In Substitute Meeting Invitation Email] The value of element email:To is set to the email address of the delegate.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5605");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5605
            // If Sync response does not contain CC element, then MS-ASCMD_R5605 is verified
            Site.CaptureRequirementIfIsNull(
                (string)TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Cc),
                5605,
                @"[In Substitute Meeting Invitation Email] The value of element email:Cc is blank.");

            string substituteInvitationEmailSubject = (string)TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Subject1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5607");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5607
            Site.CaptureRequirementIfIsTrue(
                substituteInvitationEmailSubject.Contains(originalSubject),
                5607,
                @"[In Substitute Meeting Invitation Email] The value of element email:Subject is original subject prepended with explanatory text.");
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User8Information, this.User8Information.InboxCollectionId, substituteInvitationEmailSubject);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify if delegate user forward substitute meeting invitation email, server will append the original meeting request to the forwarded message. 
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC14_MeetingResponse_ForwardSubstituteMeetingInvitationEmail()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5822, this.Site), "Exchange Server 2013 and above support Substitute Meeting Invitation.");

            #region User1 calls SendMail command to send one meeting request to user7
            string originalSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User7Information.UserName, this.User7Information.UserDomain);
            Calendar calendar = this.CreateCalendar(originalSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user7
            this.SendMeetingRequest(originalSubject, calendar);
            #endregion

            #region Get new added meeting request email in user7 mailbox
            // Switch to user7 mailbox
            this.SwitchUser(this.User7Information);

            // Sync Inbox folder
            this.GetMailItem(this.User7Information.InboxCollectionId, originalSubject);

            // Sync Calendar folder
            this.GetMailItem(this.User7Information.CalendarCollectionId, originalSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User7Information, this.User7Information.CalendarCollectionId, originalSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User7Information, this.User7Information.InboxCollectionId, originalSubject);
            #endregion

            #region Get substitute invitation email in user8 mailbox
            // Switch to user8 mailbox
            this.SwitchUser(this.User8Information);

            // Sync Inbox folder
            SyncResponse substituteSyncResponse = this.GetSubstituteMailItem(this.User8Information.InboxCollectionId, originalSubject);
            string substituteInvitationEmailServerId = TestSuiteBase.FindServerId(substituteSyncResponse, "ThreadTopic", originalSubject);
            string substituteInvitationEmailSubject = (string)TestSuiteBase.GetElementValueFromSyncResponse(substituteSyncResponse, substituteInvitationEmailServerId, Response.ItemsChoiceType8.Subject1);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User8Information, this.User8Information.InboxCollectionId, substituteInvitationEmailSubject);
            #endregion

            #region User8 creates SmartForward request which forwards mail to user2
            string forwardFromUser = Common.GetMailAddress(this.User8Information.UserName, this.User8Information.UserDomain);
            string forwardToUser = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string forwardContent = Common.GenerateResourceName(Site, "forward:body");
            SmartForwardRequest smartForwardRequest = this.CreateSmartForwardRequest(this.User8Information.InboxCollectionId, substituteInvitationEmailServerId, forwardFromUser, forwardToUser, string.Empty, string.Empty, substituteInvitationEmailSubject, forwardContent);
            #endregion

            #region User8 calls SmartForward command
            SmartForwardResponse smartForwardResponse = this.CMDAdapter.SmartForward(smartForwardRequest);
            Site.Assert.IsTrue(string.IsNullOrEmpty(smartForwardResponse.ResponseDataXML), "If SmartForward command execute success, server will return empty");
            #endregion

            #region User2 calls Sync command to get mailbox change
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);

            // Sync user2 Inbox folder
            this.GetMailItem(this.User2Information.InboxCollectionId, substituteInvitationEmailSubject);
            #endregion

            #region Record user name, folder collectionId and item subject that are generated in this case
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, substituteInvitationEmailSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, substituteInvitationEmailSubject);

            // Check the meeting forward notification mail which is sent from server to User1.
            this.SwitchUser(this.User1Information);
            string notificationSubject = "Meeting Forward Notification: " + substituteInvitationEmailSubject;
            this.CheckMeetingForwardNotification(this.User1Information, notificationSubject);
            #endregion
        }

        /// <summary>
        /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC15_MeetingResponse_AcceptMeetingInCalendarFolder()
        {
            Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(this.Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

            #region User1 calls SendMail command to send one meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);

            // Sync Inbox folder
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string serverIDForMeetingRequest = TestSuiteBase.FindServerId(syncResponse, "Subject", meetingRequestSubject);

            // Sync Calendar folder
            SyncResponse syncCalendarBeforeMeetingResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemServerID = TestSuiteBase.FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to accept the meeting request in the user2's Calendar folder.
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(1, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

            // If the user accepts the meeting request, the meeting request mail will be deleted and calendar item will be created.
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);

            bool isVerifiedR5707 = meetingResponseResponse.ResponseData.Result[0].CalendarId != null && meetingResponseResponse.ResponseData.Result[0].Status != "0";

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR5707,
                5707,
                @"[In MeetingResponse] The MeetingResponse command is used to accept [, tentatively accept, or decline] a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC16_MeetingResponse_TentativelyAcceptInCalendarFolder()
        {
            Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(this.Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

            #region User1 calls SendMail command to send one meeting request to user2.
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region User2 calls Sync command to sync user2 mailbox changes
            this.SwitchUser(this.User2Information);
            SyncResponse syncInboxResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string inboxItemID = TestSuiteBase.FindServerId(syncInboxResponse, "Subject", meetingRequestSubject);

            SyncResponse syncCalendarResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemID = TestSuiteBase.FindServerId(syncCalendarResponse, "Subject", meetingRequestSubject);

            // Get calendar item responseType value before meetingResponse
            // ResponseType is not supported in ProtocolVersion 12.1, refer MS-ASCAL <17> Section 2.2.2.38
            uint responseTypeBeforeMeetingResponse = 0;
            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.0")
            {
                responseTypeBeforeMeetingResponse = (uint)TestSuiteBase.GetElementValueFromSyncResponse(syncCalendarResponse, calendarItemID, Response.ItemsChoiceType8.ResponseType);
            }

            // Record relative items for clean up
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.CalendarCollectionId, meetingRequestSubject);
            #endregion

            #region User2 calls MeetingResponse command to tentative accept the meeting
            // Set to tentatively accept one of recurring meeting request instance
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(2, this.User2Information.InboxCollectionId, inboxItemID, null);
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);
            Site.Assert.AreEqual<int>(1, int.Parse(meetingResponseResponse.ResponseData.Result[0].Status), "If MeetingResponse command executes successfully, server should return status 1");
            string itemServerID = this.GetItemServerIdFromSpecialFolder(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, meetingRequestSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, this.User2Information.DeletedItemsCollectionId, meetingRequestSubject);

            bool isVerifiedR5708 = itemServerID != null && meetingResponseResponse.ResponseData.Result[0].Status == "1";

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR5708,
                5708,
                @"[In MeetingResponse] The MeetingResponse command is used to [accept,] tentatively accept [, or decline] a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify CalendarId is returned, when the meeting request is accepted.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S09_TC17_MeetingResponse_DeclineInCalendarFolder()
        {
            Site.Assume.AreNotEqual<SutVersion>(SutVersion.ExchangeServer2007, Common.GetSutVersion(this.Site), "Exchange 2007 SP3 does not use MeetingResponse command on Calendar folder.");

            #region User1 calls SendMail command to send one meeting request to user2
            string meetingRequestSubject = Common.GenerateResourceName(Site, "subject");
            string attendeeEmailAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            Calendar calendar = this.CreateCalendar(meetingRequestSubject, attendeeEmailAddress, null);

            // Send a meeting request email to user2
            this.SendMeetingRequest(meetingRequestSubject, calendar);
            #endregion

            #region Get new added meeting request email
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information);

            // Sync Inbox folder
            SyncResponse syncResponse = this.GetMailItem(this.User2Information.InboxCollectionId, meetingRequestSubject);
            string serverIDForMeetingRequest = TestSuiteBase.FindServerId(syncResponse, "Subject", meetingRequestSubject);

            // Sync Calendar folder
            SyncResponse syncCalendarBeforeMeetingResponse = this.GetMailItem(this.User2Information.CalendarCollectionId, meetingRequestSubject);
            string calendarItemServerID = TestSuiteBase.FindServerId(syncCalendarBeforeMeetingResponse, "Subject", meetingRequestSubject);
            #endregion

            #region Call method MeetingResponse to decline the meeting request in the user2's Calendar folder.
            MeetingResponseRequest meetingResponseRequest = TestSuiteBase.CreateMeetingResponseRequest(3, this.User2Information.InboxCollectionId, serverIDForMeetingRequest, string.Empty);

            // If the user declines the meeting request, the meeting request mail will be deleted and no calendar item will be created.
            MeetingResponseResponse meetingResponseResponse = this.CMDAdapter.MeetingResponse(meetingResponseRequest);

            SyncResponse syncInboxFolder = this.SyncChanges(this.User2Information.CalendarCollectionId);
            string itemServerIDInCalendarFolder = TestSuiteBase.FindServerId(syncInboxFolder, "Subject", meetingRequestSubject);

            bool isVerifiedR5709 = meetingResponseResponse.ResponseData.Result[0].Status == "1" && itemServerIDInCalendarFolder == null;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR5709,
                5709,
                @"[In MeetingResponse] The MeetingResponse command is used to [accept, tentatively accept, or] decline a meeting request in the [user's Inbox folder or] Calendar folder.<3>");
            #endregion
        }

        #endregion

        #region Private methods
        /// <summary>
        /// Create a MeetingResponse request.
        /// </summary>
        /// <param name="userResponse">The way the user response the meeting.</param>
        /// <param name="collectionID">The collection id of the folder that contains the meeting request.</param>
        /// <param name="requestIDList">The server ID list of the meeting request message item.</param>
        /// <param name="instanceID">The instance ID of the recurring meeting to be modified.</param>
        /// <returns>The MeetingResponse request.</returns>
        private static MeetingResponseRequest CreateMultiMeetingResponseRequest(byte userResponse, string collectionID, List<string> requestIDList, string instanceID)
        {
            List<Request.MeetingResponseRequest> requestList = new List<Request.MeetingResponseRequest>();
            foreach (string requestID in requestIDList)
            {
                Request.MeetingResponseRequest request = new Request.MeetingResponseRequest
                {
                    CollectionId = collectionID,
                    RequestId = requestID,
                    UserResponse = userResponse
                };

                // Set the instanceId of the meeting request to response
                if (!string.IsNullOrEmpty(instanceID))
                {
                    request.InstanceId = instanceID;
                }

                requestList.Add(request);
            }

            return Common.CreateMeetingResponseRequest(requestList.ToArray());
        }

        /// <summary>
        /// Get email with special threadTopic
        /// </summary>
        /// <param name="folderID">The folderID that store mail items</param>
        /// <param name="threadTopic">The thread topic</param>
        /// <returns>Sync result</returns>
        private SyncResponse GetSubstituteMailItem(string folderID, string threadTopic)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            SyncResponse syncResult = this.SyncChanges(folderID);
            string serverID = TestSuiteBase.FindServerId(syncResult, "ThreadTopic", threadTopic);
            while (serverID == null && counter < retryCount)
            {
                Thread.Sleep(waitTime);
                syncResult = this.SyncChanges(folderID);
                if (syncResult.ResponseDataXML != null)
                {
                    serverID = TestSuiteBase.FindServerId(syncResult, "ThreadTopic", threadTopic);
                }

                counter++;
            }

            Site.Assert.IsNotNull(serverID, "The email item with subject '{0}' should be found.", threadTopic);
            Site.Log.Add(LogEntryKind.Debug, "Find item successful Loop count {0}", counter);
            return syncResult;
        }

        /// <summary>
        /// Delete all items from the specified collection
        /// </summary>
        /// <param name="collectionId">The specified collection id</param>
        private void DeleteAll(string collectionId)
        {
            ItemOperationsRequest request = new ItemOperationsRequest
            {
                RequestData = new Request.ItemOperations()
                {
                    Items = new object[]
                    {
                        new Request.ItemOperationsEmptyFolderContents
                        {
                            CollectionId = collectionId,
                            Options = new Request.ItemOperationsEmptyFolderContentsOptions
                            {
                                DeleteSubFolders = string.Empty
                            }
                        }
                    }
                }
            };

            ItemOperationsResponse response = this.CMDAdapter.ItemOperations(request, DeliveryMethodForFetch.Inline);
            Site.Assert.IsTrue(response.ResponseData != null && response.ResponseData.Status == "1" && response.ResponseData.Response.EmptyFolderContents[0].Status == "1", "All items in the specified collection should be deleted successfully.");
        }
        #endregion
    }
}