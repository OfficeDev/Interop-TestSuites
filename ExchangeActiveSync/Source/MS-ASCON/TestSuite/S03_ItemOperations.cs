namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System;
    using System.Collections.ObjectModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Request;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;

    /// <summary>
    /// This scenario is designed to ignore a conversation, set up a conversation to be moved always and request a Message part using ItemOperations command.
    /// </summary>
    [TestClass]
    public class S03_ItemOperations : TestSuiteBase
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

        #region MSASCON_S03_TC01_ItemOperations_Ignore
        /// <summary>
        /// This test case is designed to validate ignoring a conversation by ItemOperations command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC01_ItemOperations_Ignore()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem sourceConversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call ItemOperations command to ignore the conversation
            // Move the created conversation from Inbox folder to Deleted Items folder with setting MoveAlways element.
            ItemOperationsResponse itemOperationResponse = this.ItemOperationsMove(sourceConversationItem.ConversationId, User1Information.DeletedItemsCollectionId, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.DeletedItemsCollectionId, conversationSubject, false);

            bool isVerifiedR214 = itemOperationResponse.ResponseData.Response.Move[0].Status != null && itemOperationResponse.ResponseData.Response.Move[0].ConversationId != null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R214");
            Site.Log.Add(LogEntryKind.Debug, "The value of the Status element is {0}, and the value of ConversationId element is {1}.", itemOperationResponse.ResponseData.Response.Move[0].Status, itemOperationResponse.ResponseData.Response.Move[0].ConversationId);

            // Verify MS-ASCON requirement: MS-ASCON_R214
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR214,
                214,
                @"[In Ignoring a Conversation] The server sends an ItemOperations command response ([MS-ASCMD] section 2.2.1.10) that includes the itemoperations:Status element, as specified in section 2.2.2.10, and the itemoperations:ConversationId element (section 2.2.2.3.1).");

            Site.Assert.AreEqual("1", itemOperationResponse.ResponseData.Response.Move[0].Status, "The move operation should be success.");
            #endregion

            #region User1 synchronizes messages in the Inbox folder and Deleted Items folder after conversation moved
            DataStructures.Sync syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);
            Site.Assert.IsNull(syncResult, "No conversation messages should not be found in Inbox folder.");

            ConversationItem destinationCoversationItem = this.GetConversationItem(User1Information.DeletedItemsCollectionId, sourceConversationItem.ConversationId);
            Site.Assert.AreEqual(sourceConversationItem.ServerId.Count, destinationCoversationItem.ServerId.Count, "All conversation messages should be moved to Deleted Items folder.");
            #endregion

            #region User2 replies the received message to create a future e-mail message for that conversation.
            this.SwitchUser(this.User2Information, false);
            syncResult = this.SyncEmail(conversationSubject, User2Information.InboxCollectionId, true, null, null);

            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);

            // Smart reply the received email from User2 to User1.
            this.CallSmartReplyCommand(syncResult.ServerId, User2Information.InboxCollectionId, user2MailboxAddress, user1MailboxAddress, conversationSubject);

            // Switch the current user to User1.
            this.SwitchUser(this.User1Information, false);

            destinationCoversationItem = this.GetConversationItem(User1Information.DeletedItemsCollectionId, sourceConversationItem.ConversationId, sourceConversationItem.ServerId.Count + 1);
            Site.Assert.AreEqual(sourceConversationItem.ServerId.Count + 1, destinationCoversationItem.ServerId.Count, "The future message should be moved to Deleted Items folder.");

            // Check if the received email is in Inbox folder.
            syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);
            Site.Assert.IsNull(syncResult, "The future message should not be found in Inbox folder.");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC02_ItemOperations_MoveAlways
        /// <summary>
        /// This test case is designed to validate always moving a conversation by ItemOperations command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC02_ItemOperations_MoveAlways()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem sourceConversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Move one email in the conversation to Sent Items folder
            Collection<string> idToMove = new Collection<string>();
            idToMove.Add(sourceConversationItem.ServerId[0]);
            this.CallMoveItemsCommand(idToMove, User1Information.InboxCollectionId, User1Information.SentItemsCollectionId);
            #endregion

            #region Call ItemOperations command to move the conversation from Inbox to DeletedItems folder
            ItemOperationsResponse itemOperationResponse = this.ItemOperationsMove(sourceConversationItem.ConversationId, User1Information.DeletedItemsCollectionId, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.SentItemsCollectionId, conversationSubject, false);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.DeletedItemsCollectionId, conversationSubject, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R143");

            // Verify MS-ASCON requirement: MS-ASCON_R143
            Site.CaptureRequirementIfAreEqual(
                "1",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                143,
                @"[In Status] [The meaning of status value] 1 [is] Success. The server successfully completed the operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R216");

            // Verify MS-ASCON requirement: MS-ASCON_R216
            // If R143 has been captured and the ConversationId element is not null, then this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemOperationResponse.ResponseData.Response.Move[0].ConversationId,
                216,
                @"[In Always Moving a Conversation] The server sends an ItemOperations command response ([MS-ASCMD] section 2.2.1.10) that includes the itemoperations:Status element, as specified in section 2.2.2.10, and the itemoperations:ConversationId element (section 2.2.2.3.1).");
            #endregion

            #region User1 syncs messages in the Inbox folder and Deleted Items folder after conversation moved
            DataStructures.Sync syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);
            Site.Assert.IsNull(syncResult, "No conversation messages should not be found in Inbox folder.");

            ConversationItem destinationCoversationItem = this.GetConversationItem(User1Information.DeletedItemsCollectionId, sourceConversationItem.ConversationId);
            Site.Assert.AreEqual(sourceConversationItem.ServerId.Count - 1, destinationCoversationItem.ServerId.Count, "All conversation messages except in Sent Items Folder should be moved to Deleted Items folder.");
            #endregion

            #region User2 replies the received message to create a future e-mail message for that conversation.
            this.SwitchUser(this.User2Information, false);
            syncResult = this.SyncEmail(conversationSubject, User2Information.InboxCollectionId, true, null, null);

            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);

            // Smart reply the received email from User2 to User1.
            this.CallSmartReplyCommand(syncResult.ServerId, User2Information.InboxCollectionId, user2MailboxAddress, user1MailboxAddress, conversationSubject);

            // Switch the current user to User1.
            this.SwitchUser(this.User1Information, false);

            destinationCoversationItem = this.GetConversationItem(User1Information.DeletedItemsCollectionId, sourceConversationItem.ConversationId, sourceConversationItem.ServerId.Count + 1);
            Site.Assert.AreEqual(sourceConversationItem.ServerId.Count, destinationCoversationItem.ServerId.Count, "The future message should be moved to Deleted Items folder.");

            // Check if the received email is in Inbox folder.
            syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);
            Site.Assert.IsNull(syncResult, "The future message should not be found in Inbox folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R68");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                68,
                @"[In ConversationId (ItemOperations)] In an ItemOperations command request ([MS-ASCMD] section 2.2.1.10), the itemoperations:ConversationId element ([MS-ASCMD] section 2.2.3.35.1) is a required child element of the itemoperations:Move element ([MS-ASCMD] section 2.2.3.117.1) that specifies the conversation ID of the conversation that is to be moved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R118");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                118,
                @"[In DstFldId] The itemoperations:DstFldId element ([MS-ASCMD] section 2.2.3.51.1) is a required child element of the itemoperations:Move element ([MS-ASCMD] section 2.2.3.117.1) in an ItemOperations command request ([MS-ASCMD] section 2.2.1.10) that specifies the folder to which the conversation is moved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R130");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                130,
                @"[In MoveAlways] When a conversation is set to always be moved, all e-mail messages in the conversation, including all future e-mail messages in the conversation, are moved from all folders except the Sent Items folder to the destination folder that is specified by the DstFldId element (section 2.2.2.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R166");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                166,
                @"[In Ignoring a Conversation] When a conversation is ignored, all e-mail messages in the conversation, including all future e-mail messages for that conversation, are moved from all folders except Sent Items folder to the Deleted Items folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R172");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                172,
                @"[In Setting up a Conversation to Be Moved Always] When a conversation is set to be moved always, all e-mail messages in the conversation, including all future e-mail messages for that conversation, are moved from all folders except Sent Items folder to a destination folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R213");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                213,
                @"[In Ignoring a Conversation] When the server receives a request to ignore a conversation, as specified in section 3.1.4.4, the server moves all e-mail messages in the conversation, including all future e-mail messages for that conversation, from all folders except Sent Items folder to the Deleted Items folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R215");

            // If all messages except the one in Sent Items Folder and future message are moved to DeletedItems folder, then this requirement can be captured.
            Site.CaptureRequirement(
                215,
                @"[In Always Moving a Conversation] When the server receives a request to always move a conversation, as specified in section 3.1.4.6, the server moves all e-mail messages in the conversation, including all future e-mail messages for that conversation, from all folders except Sent Items folder to the specified destination folder.");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC03_ItemOperations_Status2
        /// <summary>
        /// This test case is designed to validate status 2 is returned for ItemOperations command if the ItemOperations command request is invalid.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC03_ItemOperations_Status2()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call ItemOperations command with invalid MoveAlways element.
            // Create an ItemOperations request and set invalid content to the MoveAlways element.
            ItemOperationsMove move = new ItemOperationsMove
            {
                DstFldId = User1Information.DeletedItemsCollectionId,
                ConversationId = conversationItem.ConversationId,
                Options = new ItemOperationsMoveOptions { MoveAlways = "MoveAlwaysContent" }
            };

            ItemOperationsRequest itemOperationRequest = Common.CreateItemOperationsRequest(new object[] { move });
            ItemOperationsResponse itemOperationResponse = this.CONAdapter.ItemOperations(itemOperationRequest);
            Site.Assert.AreEqual("1", itemOperationResponse.ResponseData.Status, "The ItemOperations operation should be success.");
            Site.Assert.AreEqual(1, itemOperationResponse.ResponseData.Response.Move.Length, "The server should return a Move element in ItemOperationsResponse.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R144");

            // Verify MS-ASCON requirement: MS-ASCON_R144
            Site.CaptureRequirementIfAreEqual(
                "2",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                144,
                @"[In Status] [The meaning of status value] 2  [is] Protocol error. The XML is not valid.");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC04_ItemOperations_Status6
        /// <summary>
        /// This test case is designed to validate status 6 is returned for ItemOperations command if the conversation or destination folder does not exist.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC04_ItemOperations_Status6()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call ItemOperations command with an invalid ConversationId.
            string conversationId = Convert.ToBase64String(Encoding.Default.GetBytes("NotExistConversationId"));
            ItemOperationsResponse itemOperationResponse = this.ItemOperationsMove(conversationId, User1Information.SentItemsCollectionId, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R146");

            // Verify MS-ASCON requirement: MS-ASCON_R146
            Site.CaptureRequirementIfAreEqual(
                "6",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                146,
                @"[In Status] [The meaning of status value] 6 [is] Not Found. The conversation [or destination folder] does not exist.");
            #endregion

            #region Call ItemOperations command with an invalid destination folder id.
            // Create an ItemOperations request and set an invalid destination folder id.
            itemOperationResponse = this.ItemOperationsMove(conversationItem.ConversationId, "NotExistFolderId", true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R317");

            // Verify MS-ASCON requirement: MS-ASCON_R317
            Site.CaptureRequirementIfAreEqual(
                "6",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                317,
                @"[In Status] [The meaning of status value 6 is] Not Found. The [conversation or] destination folder does not exist.");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC05_ItemOperations_Status105
        /// <summary>
        /// This test case is designed to validate status 105 is returned for ItemOperations command if the destination folder is recipient information cache.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC05_ItemOperations_Status105()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call ItemOperations command and set DstFldId to the recipient information cache.
            ItemOperationsResponse itemOperationResponse = this.ItemOperationsMove(conversationItem.ConversationId, User1Information.RecipientInformationCacheCollectionId, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R262");

            // Verify MS-ASCON requirement: MS-ASCON_R262
            Site.CaptureRequirementIfAreEqual(
                "105",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                262,
                @"[In Status] [The meaning of status value] 105 [is] Invalid Combination of IDs. The destination folder cannot be the Recipient Information Cache.");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC06_ItemOperations_Status156
        /// <summary>
        /// This test case is designed to validate status 156 is returned for ItemOperations command if the destination folder is not of type IPF.Note.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC06_ItemOperations_Status156()
        {
            #region Create a conversation and get the created conversation item
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call ItemOperations command and set DstFldId to the Calendar folder.
            // Task folder is not the IPF.NOTE type.
            ItemOperationsResponse itemOperationResponse = this.ItemOperationsMove(conversationItem.ConversationId, User1Information.CalendarCollectionId, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R338");

            // Verify MS-ASCON requirement: MS-ASCON_R338
            Site.CaptureRequirementIfAreEqual(
                "156",
                itemOperationResponse.ResponseData.Response.Move[0].Status,
                338,
                @"[In Status] [The meaning of status value] 156 [is] Action not supported. The destination folder MUST be of type ""IPF.Note"".");
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC07_ItemOperations_Status164
        /// <summary>
        /// This test case is designed to validate status 164 is returned if a value other than 2 is specified in the Type element of BodyPartPreference element in ItemOperations command request.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC07_ItemOperations_Status164()
        {
            this.CheckActiveSyncVersionIsNot140();

            #region User2 sends an email to User1
            this.SwitchUser(this.User2Information, true);

            string subject = Common.GenerateResourceName(Site, "Subject");
            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
           
            this.CallSendMailCommand(user2MailboxAddress, user1MailboxAddress, subject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, subject, false);
            #endregion

            #region Call ItemOperations command with BodyPartPreference element and set the Type element to 3
            this.SwitchUser(this.User1Information, false);

            DataStructures.Sync syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, null, null);
            BodyPartPreference bodyPartPreference = new BodyPartPreference()
            {
                Type = 3,
            };

            ItemOperationsRequest itemOperationsRequest = TestSuiteHelper.GetItemOperationsRequest(User1Information.InboxCollectionId, syncItem.ServerId, bodyPartPreference, null);
            ItemOperationsResponse itemOperationsResponse = this.CONAdapter.ItemOperations(itemOperationsRequest);
            Site.Assert.AreEqual("1", itemOperationsResponse.ResponseData.Status, "The ItemOperations operation should be success.");
            this.VerifyMessagePartStatus164(byte.Parse(itemOperationsResponse.ResponseData.Response.Fetch[0].Status));
            #endregion
        }
        #endregion

        #region MSASCON_S03_TC08_ItemOperations_MessagePart
        /// <summary>
        /// This test case is designed to validate requesting the message part by ItemOperations command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S03_TC08_ItemOperations_MessagePart()
        {
            this.CheckActiveSyncVersionIsNot140();

            #region User2 sends an email to User1
            this.SwitchUser(this.User2Information, true);

            string subject = Common.GenerateResourceName(Site, "Subject");
            string body = Common.GenerateResourceName(Site, "Body");
            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
            this.CallSendMailCommand(user2MailboxAddress, user1MailboxAddress, subject, body);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, subject, false);
            #endregion

            #region Call ItemOperations command without BodyPreference or BodyPartPreference element.
            this.SwitchUser(this.User1Information, false);

            // Get all of the email BodyPart data.
            BodyPartPreference bodyPartPreference = new BodyPartPreference()
            {
                Type = 2,
            };

            DataStructures.Sync syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, bodyPartPreference, null);
            XmlElement lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
            string allData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);

            DataStructures.Email email = this.ItemOperationsFetch(User1Information.InboxCollectionId, syncItem.ServerId, null, null);
            this.VerifyMessagePartWithoutPreference(email);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R340");

            // Verify MS-ASCON requirement: MS-ASCON_R340
            Site.CaptureRequirementIfIsNull(
                email.BodyPart,
                340,
                @"[In Sending a Message Part] The airsyncbase:BodyPart element is not present in the [ItemOperations command] response if the client did not request the message part, as specified in section 3.1.4.10.");
            #endregion

            #region Call ItemOperations command with only BodyPreference element.
            BodyPreference bodyPreference = new BodyPreference()
            {
                Type = 2,
            };

            email = this.ItemOperationsFetch(User1Information.InboxCollectionId, syncItem.ServerId, null, bodyPreference);
            this.VerifyMessagePartWithBodyPreference(email);
            #endregion

            #region Call ItemOperations command with only BodyPartPreference element.
            bodyPartPreference = new BodyPartPreference()
            {
                Type = 2,
                TruncationSize = 12,
                TruncationSizeSpecified = true,
            };

            email = this.ItemOperationsFetch(User1Information.InboxCollectionId, syncItem.ServerId, bodyPartPreference, null);
            lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
            string truncatedData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);
            this.VerifyMessagePartWithBodyPartPreference(email, truncatedData, allData, (int)bodyPartPreference.TruncationSize);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R236");

            // Verify MS-ASCON requirement: MS-ASCON_R236
            Site.CaptureRequirementIfIsNotNull(
                email.BodyPart,
                236,
                @"[In Sending a Message Part] If the client [Sync command request ([MS-ASCMD] section 2.2.1.21), Search command request ([MS-ASCMD] section 2.2.1.16) or] ItemOperations command request 9([MS-ASCMD] section 2.2.1.10) includes the airsyncbase:BodyPartPreference element (section 2.2.2.2), then the server uses the airsyncbase:BodyPart element (section 2.2.2.1) to encapsulate the message part in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R39");

            // A message part and its meta-data are encapsulated by BodyPart element in the ItemOperation response, so this requirement can be captured.
            Site.CaptureRequirement(
                39,
                @"[In BodyPart] The airsyncbase:BodyPart element ([MS-ASAIRS] section 2.2.2.10) encapsulates a message part and its meta-data in [a Sync command response ([MS-ASCMD] section 2.2.1.21),] an ItemOperations command response ([MS-ASCMD] section 2.2.1.10) [or a Search command response ([MS-ASCMD] section 2.2.1.16)].");
            #endregion

            #region Call ItemOperations command with both BodyPreference and BodyPartPreference elements.
            email = this.ItemOperationsFetch(User1Information.InboxCollectionId, syncItem.ServerId, bodyPartPreference, bodyPreference);
            this.VerifyMessagePartWithBothPreference(email);
            #endregion
        }
        #endregion
    }
}