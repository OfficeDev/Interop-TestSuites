namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System.Collections.ObjectModel;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Request;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;

    /// <summary>
    /// This scenario is designed to find a conversation using Search command.
    /// </summary>
    [TestClass]
    public class S05_Search : TestSuiteBase
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

        #region MSASCON_S05_TC01_Search
        /// <summary>
        /// This test case is designed to validate searching a conversation by Search command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S05_TC01_Search()
        {
            if (Common.IsRequirementEnabled(221, this.Site))
            {               
                #region Create a conversation and get the created conversation item.
                string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
                ConversationItem sourceConversationItem = this.CreateConversation(conversationSubject);
                #endregion

                #region Call MoveItems command to move one item of the conversation from Inbox folder to SentItems folder.
                Collection<string> moveItems = new Collection<string> { sourceConversationItem.ServerId[0] };

                // Call MoveItems command to move one item of the conversation from Inbox folder to SentItems folder.
                this.CallMoveItemsCommand(moveItems, User1Information.InboxCollectionId, User1Information.SentItemsCollectionId);
                TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.SentItemsCollectionId, conversationSubject, false);
                #endregion

                #region Call Search command to find the conversation.
                DataStructures.SearchStore searchResponse = this.CallSearchCommand(sourceConversationItem.ConversationId, 2, null, null);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R221");

                // Verify MS-ASCON requirement: MS-ASCON_R221
                // The Search command executed successfully, so this requirement can be captured.
                Site.CaptureRequirement(
                    221,
                    @"[In Processing a Search Command] The server sends a Search command response, as specified in [MS-ASCMD] section 2.2.1.16.");

                Site.Assert.AreEqual<int>(searchResponse.Results.Count, sourceConversationItem.ServerId.Count, "The count of the search result should be equal to the count of items in the conversation.");

                // If one of the found email is in Inbox folder and the other is in Sent Items folder, this requirement can be captured.
                bool allFoldersSearched = (searchResponse.Results[0].CollectionId == User1Information.InboxCollectionId && searchResponse.Results[1].CollectionId == User1Information.SentItemsCollectionId) || (searchResponse.Results[1].CollectionId == User1Information.InboxCollectionId && searchResponse.Results[0].CollectionId == User1Information.SentItemsCollectionId);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R176");
                Site.Log.Add(LogEntryKind.Debug, "The emails found are in folders with CollectionId {0} and {1}.", searchResponse.Results[0].CollectionId, searchResponse.Results[1].CollectionId);

                // Verify MS-ASCON requirement: MS-ASCON_R176
                Site.CaptureRequirementIfIsTrue(
                    allFoldersSearched,
                    176,
                    @"[In Finding a Conversation] Searching for a particular conversation will search across all folders for all e-mail messages that are in the conversation.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R220");
                Site.Log.Add(LogEntryKind.Debug, "The emails found are in folders with CollectionId {0} and {1}.", searchResponse.Results[0].CollectionId, searchResponse.Results[1].CollectionId);

                // Verify MS-ASCON requirement: MS-ASCON_R220
                Site.CaptureRequirementIfIsTrue(
                    allFoldersSearched,
                    220,
                     @"[In Processing a Search Command] When the server receives a request to find a conversation, as specified in section 3.1.4.7, the server searches across all folders for all e-mail messages that are in the conversation and returns this set of e-mail messages.");
                #endregion
            }
        }
        #endregion

        #region MSASCON_S05_TC02_Search_MessagePart
        /// <summary>
        /// This test case is designed to validate requesting the message part by Search command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S05_TC02_Search_MessagePart()
        {
            if (Common.IsRequirementEnabled(221, this.Site))
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

                #region Call Search command without BodyPreference or BodyPartPreference element.
                this.SwitchUser(this.User1Information, false);

                // Get all of the email BodyPart data.
                BodyPartPreference bodyPartPreference = new BodyPartPreference()
                {
                    Type = 2,
                };

                DataStructures.Sync syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, bodyPartPreference, null);
                XmlElement lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
                string allData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);

                DataStructures.SearchStore searchStore = this.CallSearchCommand(syncItem.Email.ConversationId, 1, null, null);
                this.VerifyMessagePartWithoutPreference(searchStore.Results[0].Email);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R339");

                // Verify MS-ASCON requirement: MS-ASCON_R339
                Site.CaptureRequirementIfIsNull(
                    searchStore.Results[0].Email.BodyPart,
                    339,
                    @"[In Sending a Message Part] The airsyncbase:BodyPart element is not present in the [Search command] response if the client did not request the message part, as specified in section 3.1.4.10.");
                #endregion

                #region Call Search command with BodyPreference element.
                BodyPreference bodyPreference = new BodyPreference()
                {
                    Type = 2,
                };

                searchStore = this.CallSearchCommand(syncItem.Email.ConversationId, 1, null, bodyPreference);
                this.VerifyMessagePartWithBodyPreference(searchStore.Results[0].Email);
                #endregion

                #region Call Search command with BodyPartPreference element.
                bodyPartPreference = new BodyPartPreference()
                {
                    Type = 2,
                    TruncationSize = 12,
                    TruncationSizeSpecified = true,
                };

                searchStore = this.CallSearchCommand(syncItem.Email.ConversationId, 1, bodyPartPreference, null);
                lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
                string truncatedData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);
                this.VerifyMessagePartWithBodyPartPreference(searchStore.Results[0].Email, truncatedData, allData, (int)bodyPartPreference.TruncationSize);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R235");

                // Verify MS-ASCON requirement: MS-ASCON_R235
                Site.CaptureRequirementIfIsNotNull(
                    searchStore.Results[0].Email.BodyPart,
                    235,
                    @"[In Sending a Message Part] If the client [Sync command request ([MS-ASCMD] section 2.2.1.21),] Search command request ([MS-ASCMD] section 2.2.1.16) [or ItemOperations command request 9([MS-ASCMD] section 2.2.1.10)] includes the airsyncbase:BodyPartPreference element(section 2.2.2.2), then the server uses the airsyncbase:BodyPart element (section 2.2.2.1) to encapsulate the message part in the response.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R40");

                // A message part and its meta-data are encapsulated by BodyPart element in the Search response, so this requirement can be captured.
                Site.CaptureRequirement(
                    40,
                    @"[In BodyPart] The airsyncbase:BodyPart element ([MS-ASAIRS] section 2.2.2.10) encapsulates a message part and its meta-data in [a Sync command response ([MS-ASCMD] section 2.2.1.21), an ItemOperations command response ([MS-ASCMD] section 2.2.1.10) or] a Search command response ([MS-ASCMD] section 2.2.1.16).");
                #endregion

                #region Call Search command with both BodyPreference and BodyPartPreference elements.
                searchStore = this.CallSearchCommand(syncItem.Email.ConversationId, 1, bodyPartPreference, bodyPreference);
                this.VerifyMessagePartWithBothPreference(searchStore.Results[0].Email);
                #endregion
            }
        }
        #endregion

        #region MSASCON_S05_TC03_Search_Status164
        /// <summary>
        /// This test case is designed to validate Status 164 is returned if a value other than 2 is specified in the Type element of BodyPartPreference element in Search command request.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S05_TC03_Search_Status164()
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

            #region Call Search command with BodyPartPreference element and set the Type element to 3
            this.SwitchUser(this.User1Information, false);

            DataStructures.Sync syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, null, null);
            BodyPartPreference bodyPartPreference = new BodyPartPreference()
            {
                Type = 3,
            };

            SearchRequest searchRequest = TestSuiteHelper.GetSearchRequest(syncItem.Email.ConversationId, bodyPartPreference, null);
            DataStructures.SearchStore searchStore = this.CONAdapter.Search(searchRequest, false, 0);
            this.VerifyMessagePartStatus164(byte.Parse(searchStore.StoreStatus));
            #endregion
        }
        #endregion
    }
}