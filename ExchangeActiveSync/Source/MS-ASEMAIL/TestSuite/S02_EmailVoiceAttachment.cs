namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test voice attachment e-mail events, including sending an e-mail with voice attachment to server, synchronizing the e-mail with voice attachment with server.
    /// </summary>
    [TestClass]
    public class S02_EmailVoiceAttachment : TestSuiteBase
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

        #region MSASEMAIL_S02_TC01_VoiceAttachment_VerifyAllRelativeElements
        /// <summary>
        /// This case is designed to test the UmAttDuration element, UmUserNotes element, UmAttOrder element and UmCallerID element in an email voice attachment.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S02_TC01_VoiceAttachment_VerifyAllRelativeElements()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The UmAttDuration element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call SendMail command to send a mail with an electronic voice mail attachment
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string firstVoiceFilePath = "testVoice.mp3";
            string secondVoiceFilePath = "secondTestVoice.mp3";
            Sync mailResult = this.SendVoiceMail(emailSubject, firstVoiceFilePath, secondVoiceFilePath);
            #endregion

            #region Verify requirements
            foreach (Response.AttachmentsAttachment attachment in mailResult.Email.Attachments.Items)
            {
                Site.Assert.IsNotNull(attachment, "The attachment should not be null.");
                Site.Assert.IsNotNull(attachment.UmAttOrder, "The order of electronic voice mail attachment should not be null.");
            }

            // According to MS-OXCMAIL and MS-OXOUM, the order of attachments is the reverse of the order in which the attachments were added, so the most recent attachment is Email.Attachments.Attachment[1].
            Site.Assert.IsNull(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[0]).UmAttDuration, "The duration of the next recent electronic voice mail attachment should be null since this element specifies the duration of the most recent electronic voice mail attachment.");
            Site.Assert.IsNotNull(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[1]).UmAttDuration, "The duration of the most recent electronic voice mail attachment should not be null.");

            // If the server response contains Attachment element and UmAttDuration is a child element of the most recent attachment, then MS-ASEMAIL_R799 can be verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R799");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R799
            Site.CaptureRequirement(
                799,
                @"[In UmAttDuration] The email2:UmAttDuration element is an optional child element of the airsyncbase:Attachment element (section 2.2.2.7) that specifies the duration of the most recent electronic voice mail attachment in seconds.");

            // If the server response contains UmAttDuration and UmAttOrder element, the response must contains a MessageClass element which element value that begins with the prefix of "IPM.Note.Microsoft.Voicemail", "IPM.Note.RPMSG.Microsoft.Voicemail", or "IPM.Note.Microsoft.Missed.Voice", then MS-ASEMAIL_R804 and MS-ASEMAIL_R811 can be verified.
            string messageClass = mailResult.Email.MessageClass;
            bool verifyR804AndR811 = messageClass.StartsWith("IPM.Note.Microsoft.Voicemail", StringComparison.CurrentCulture) || messageClass.StartsWith("IPM.Note.RPMSG.Microsoft.Voicemail", StringComparison.CurrentCulture) || messageClass.StartsWith("IPM.Note.Microsoft.Missed.Voice", StringComparison.CurrentCulture);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R804");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R804
            Site.CaptureRequirementIfIsTrue(
                verifyR804AndR811,
                804,
                @"[In UmAttDuration] This element[email2:UmAttDuration] MUST only be included for messages with a MessageClass element (section 2.2.2.49) value that begins with the prefix of ""IPM.Note.Microsoft.Voicemail"", ""IPM.Note.RPMSG.Microsoft.Voicemail"", or ""IPM.Note.Microsoft.Missed.Voice"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R811");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R811
            Site.CaptureRequirementIfIsTrue(
                verifyR804AndR811,
                811,
                @"[In UmAttOrder] This element[email2:UmAttOrder] MUST only be included for messages with a MessageClass element (section 2.2.2.49) value that begins with the prefix of ""IPM.Note.Microsoft.Voicemail"", ""IPM.Note.RPMSG.Microsoft.Voicemail"", or ""IPM.Note.Microsoft.Missed.Voice"".");

            // If UmAttDuration element is not null means server set the element value then MS-ASEMAIL_R803 is verified.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R803");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R803
            Site.CaptureRequirementIfIsNotNull(
                ((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[1]).UmAttDuration,
                803,
                @"[In UmAttDuration] This value is set by the server [and is read-only for the client].");

            // The display name of the attachment should not be null
            Site.Assert.IsNotNull(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[0]).DisplayName, "The name of the attachment file should not be null.");

            // If the value of UmAttOrder for the most recent attachment is 1, and for the next recent is larger than the most recent one, it indicates the UmAttOrder identifies the order of electronic voice mail attachments, then MS-ASEMAIL_R805 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R805");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R805
            Site.CaptureRequirementIfIsTrue(
                int.Parse(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[1]).UmAttOrder) == 1 && int.Parse(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[0]).UmAttOrder) > 1,
                805,
                @"[In UmAttOrder] The email2:UmAttOrder element  is an optional child element of the airsyncbase:Attachment element (section 2.2.2.7) that identifies the order of electronic voice mail attachments.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R808");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R808
            Site.CaptureRequirementIfAreEqual<string>(
                "1",
                ((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[1]).UmAttOrder,
                808,
                @"[In UmAttOrder] This value is set by the server [and is read-only for the client].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R809");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R809
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                int.Parse(((Response.AttachmentsAttachment)mailResult.Email.Attachments.Items[1]).UmAttOrder),
                809,
                @"[In UmAttOrder] The most recent voice mail attachment in an e-mail item MUST have an email2:UmAttOrder value of 1.");

            if (Common.IsRequirementEnabled(829, this.Site))
            {
                // If UmCallerID element is not null, then MS-ASEMAIL_R829 can be verified
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R829");

                // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R829
                Site.CaptureRequirementIfIsNotNull(
                    mailResult.Email.UmCallerID,
                    829,
                    @"[In Appendix B: Product Behavior] Implementation does send this element[email2:UmCallerID] to the client regardless of the client's current VoIP capabilities, in order to enable future VoIP scenarios. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            // If UmCallerID element is not null, that means UmCallerID is a telephone number and is sent from the server, then MS-ASEMAIL_R812, MS-ASEMAIL_933 can be verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R812");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R812
            Site.CaptureRequirementIfIsNotNull(
                mailResult.Email.UmCallerID,
                812,
                @"[In UmCallerID] The email2:UmCallerID element is an optional element that specifies the callback telephone number of the person who called or left an electronic voice message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R933");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R933
            Site.CaptureRequirementIfIsNotNull(
                mailResult.Email.UmCallerID,
                933,
                @"[In UmCallerID] This element[UmCallID] is sent from the server to the client.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S02_TC02_VoiceAttachment_IncludedUmCallerIDInRequest
        /// <summary>
        /// This case is designed to test that if client attempts to send SyncRequest with UmCallerID element to the server, the server will return a Status element of 6 in Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S02_TC02_VoiceAttachment_IncludedUmCallerIDInRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The UmCallerID element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call SendMail command to send a mail with an electronic voice mail attachment
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string firstVoiceFilePath = "testVoice.mp3";
            string secondVoiceFilePath = "secondTestVoice.mp3";
            this.SendVoiceMail(emailSubject, firstVoiceFilePath, secondVoiceFilePath);
            #endregion

            #region Calls Sync command to update email
            string newUmCallerId = "77777777777";
            string umcallerIDElement = "<UmCallerID xmlns=\"Email2\">" + newUmCallerId + "</UmCallerID>";
            string statusCode = this.UpdateVoiceEmailWithInvalidData(umcallerIDElement);
            #endregion

            #region Verify requirement
            // If statusCode equals "6", means server response error status code 6, then MS-ASEMAIL_R817 is verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R817");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R817
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                statusCode,
                817,
                @"[In UmCallerID] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if the client attempts to send the email2:UmCallerId element to the server.");
            #endregion
        }
        #endregion

        #region MSASEMAIL_S02_TC03_VoiceAttachment_IncludedUmUserNotesInRequest
        /// <summary>
        /// This case is designed to test that if client attempts to change voice email with UmUserNotes element, the server the server will return a Status element value of 6 in the Sync command response.
        /// </summary>
        [TestCategory("MSASEMAIL"), TestMethod()]
        public void MSASEMAIL_S02_TC03_VoiceAttachment_IncludedUmUserNotesInRequest()
        {
            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The UmUserNotes element is not supported when the ActiveSyncProtocolVersion is 12.1.");

            #region Call SendMail command to send a mail with an electronic voice mail attachment
            string emailSubject = Common.GenerateResourceName(Site, "subject");
            string firstVoiceFilePath = "testVoice.mp3";
            string secondVoiceFilePath = "secondTestVoice.mp3";
            this.SendVoiceMail(emailSubject, firstVoiceFilePath, secondVoiceFilePath);
            #endregion

            #region Calls Sync command to update email
            string umuserNotes = Common.GenerateResourceName(Site, "subject");
            string umuserNotesElement = "<UmUserNotes xmlns=\"Email2\" >" + umuserNotes + "</UmUserNotes>";
            string statusCode = this.UpdateVoiceEmailWithInvalidData(umuserNotesElement);
            #endregion

            #region Verify requirement
            // If statusCode equals "6", which means server returns a status element value of 6 in response, the MS-ASEMAIL_R835 will be verified
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASEMAIL_R835");

            // Verify MS-ASEMAIL requirement: MS-ASEMAIL_R835
            Site.CaptureRequirementIfAreEqual<string>(
                "6",
                statusCode,
                835,
                @"[In UmUserNotes] The server returns a Status element ([MS-ASCMD] section 2.2.3.177.17) value of 6 in the Sync command response ([MS-ASCMD] section 2.2.1.21) if the client attempts to send the email2:UmUserNotes element to the server.");
            #endregion
        }
        #endregion

        #region Private methods
        #region Send voice mail
        /// <summary>
        /// Call SendMail command to send one voice email
        /// </summary>
        /// <param name="emailSubject">Email subject</param>
        /// <param name="firstVoiceFilePath">First voice attachment file name</param>
        /// <param name="secondVoiceFilePath">Second voice attachment file name</param>
        /// <returns>Email item</returns>
        private Sync SendVoiceMail(string emailSubject, string firstVoiceFilePath, string secondVoiceFilePath)
        {
            // Create mail content
            string senderEmail = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string receiverEmail = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            string emailBody = Common.GenerateResourceName(Site, "content");
            string callNumber = "7125550123";

            // Create voice mail content mime
            string voiceMailMime = TestSuiteHelper.CreateVoiceAttachmentMime(
                senderEmail,
                receiverEmail,
                emailSubject,
                emailBody,
                callNumber,
                firstVoiceFilePath,
                secondVoiceFilePath);

            string clientId = TestSuiteHelper.GetClientId();
            SendMailRequest sendMailRequest = TestSuiteHelper.CreateSendMailRequest(clientId, false, voiceMailMime);
            SendMailResponse response = this.EMAILAdapter.SendMail(sendMailRequest);

            // Verify send voice mail success
            Site.Assert.AreEqual<string>(
                 string.Empty,
                 response.ResponseDataXML,
                 "The server should return an empty xml response data to indicate SendMail command executes successfully.",
                 response.ResponseDataXML);

            #region Record user name, folder collectionId and item subject that are used in this case.
            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, emailSubject);
            #endregion

            #region Sync changes in user2 mailbox .
            // Sync changes
            SyncStore result = this.GetSyncResult(emailSubject, this.User2Information.InboxCollectionId, null);
            Sync emailItem = TestSuiteHelper.GetSyncAddItem(result, emailSubject);
            #endregion

            return emailItem;
        }
        #endregion

        #region Update email with UmCallerID
        /// <summary>
        /// Update email with invalid data
        /// </summary>
        /// <param name="invalidElement">invalid element send to server</param>
        /// <returns>Update results status code</returns>
        private string UpdateVoiceEmailWithInvalidData(string invalidElement)
        {
            // Switch to user2 mailbox
            this.SwitchUser(this.User2Information, true);

            // Sync changes
            SyncStore initSyncResult = this.InitializeSync(this.User2Information.InboxCollectionId);
            SyncStore syncChangeResult = this.SyncChanges(initSyncResult.SyncKey, this.User2Information.InboxCollectionId, null);
            string syncKey = syncChangeResult.SyncKey;
            string serverId = this.User2Information.InboxCollectionId;

            // Create normal Sync change request
            Request.SyncCollectionChange changeData = TestSuiteHelper.CreateSyncChangeData(true, serverId, null, null);
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncChangeRequest(syncKey, this.User2Information.InboxCollectionId, changeData);

            // Calls Sync command to update email with invalid sync request
            string insertTag = "</ApplicationData>";
            SendStringResponse result = this.EMAILAdapter.InvalidSync(syncRequest, invalidElement, insertTag);

            // Get status code
            return TestSuiteHelper.GetStatusCode(result.ResponseDataXML);
        }
        #endregion
        #endregion
    }
}