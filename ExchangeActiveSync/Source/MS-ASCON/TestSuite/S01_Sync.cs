namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System;
    using System.Collections.ObjectModel;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// This scenario is designed to mark a conversation as Read or Unread, flag a conversation for follow-up, apply a conversation-based filter, delete a conversation and request a Message part using Sync command.
    /// </summary>
    [TestClass]
    public class S01_Sync : TestSuiteBase
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

        #region MSASCON_S01_TC01_Sync_MarkRead
        /// <summary>
        /// This test case is designed to validate marking a conversation as Read or Unread by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC01_Sync_MarkRead()
        {
            #region Create a conversation and sync to get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Mark the created conversation as read.
            this.SyncChange(this.LatestSyncKey, conversationItem.ServerId, User1Information.InboxCollectionId, true, null);

            // Check whether the read property is changed to read for all messages in the conversation.
            SyncStore syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            bool markRead = true;
            int itemCount = 0;

            foreach (Sync item in syncItems.AddElements)
            {
                if (conversationItem.ServerId.Contains(item.ServerId))
                {
                    itemCount++;
                    if (item.Email.Read != true)
                    {
                        markRead = false;
                        break;
                    }
                }
            }

            Site.Assert.AreEqual<int>(2, itemCount, "The Sync response should have 2 emails belong to the conversation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R203");

            // Verify MS-ASCON requirement: MS-ASCON_R203
            // If all emails in the created conversation are marked as read, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                markRead,
                203,
                @"[In Marking a Conversation as Read or Unread] When the server receives a request to mark a conversation as read [or unread], as specified in section 3.1.4.3, the server marks all e-mails that are in the conversation and that are in the current folder as  read [or unread], whichever is specified in the client request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R326");

            // Verify MS-ASCON requirement: MS-ASCON_R326
            // If all emails in the created conversation are marked as read, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                markRead,
                326,
                @"[In Marking a Conversation as Read or Unread] When a conversation is marked as read [or unread], all e-mail messages that are in the conversation and that are in the current folder are marked as such [read].");
            #endregion

            #region Mark the created conversation as unread.
            // Mark unread for the created conversation.
            SyncStore syncStore = this.SyncChange(this.LatestSyncKey, conversationItem.ServerId, User1Information.InboxCollectionId, false, null);

            // Check whether read property is changed to unread for all messages in the conversation.
            syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            bool markUnread = true;
            itemCount = 0;

            foreach (Sync item in syncItems.AddElements)
            {
                if (conversationItem.ServerId.Contains(item.ServerId))
                {
                    itemCount++;
                    if (item.Email.Read != false)
                    {
                        markUnread = false;
                        break;
                    }
                }
            }

            Site.Assert.AreEqual<int>(2, itemCount, "The Sync response should have 2 emails belong to the conversation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R284");

            // Verify MS-ASCON requirement: MS-ASCON_R284
            // If all emails in the created conversation are marked as unread, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                markUnread,
                284,
                @"[In Marking a Conversation as Read or Unread] When the server receives a request to mark a conversation as [read or] unread  as specified in section 3.1.4.3, the server marks all e-mails that are in the conversation and that are in the current folder as [read or] unread, whichever is specified in the client's request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R337");

            // Verify MS-ASCON requirement: MS-ASCON_R337
            // If all emails in the created conversation are marked as unread, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                markUnread,
                337,
                @"[In Marking a Conversation as Read or Unread] When a conversation is marked as [read or] unread, all e-mail messages that are in the conversation and that are in the current folder are marked as such [unread].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R204");

            // Verify MS-ASCON requirement: MS-ASCON_R204
            Site.CaptureRequirementIfIsNotNull(
                syncStore,
                204,
                @"[In Marking a Conversation as Read or Unread] The server sends a Sync command response as specified in [MS-ASCMD] section 2.2.1.21.");
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC02_Sync_Flag
        /// <summary>
        /// This test case is designed to validate flagging a conversation for Follow-up by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC02_Sync_Flag()
        {
            #region Create a conversation and sync to get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem conversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Set active flag to the most recent email of the conversation.
            SyncStore syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            // Get the most recent email in the conversation.
            string mostRecentEmailServerId = null;
            Email mostRecentEmail = new Email { ConversationIndex = string.Empty };
            foreach (Sync syncItem in syncItems.AddElements)
            {
                if (conversationItem.ServerId.Contains(syncItem.ServerId) && Convert.FromBase64String(syncItem.Email.ConversationIndex).Length > Convert.FromBase64String(mostRecentEmail.ConversationIndex).Length)
                {
                    mostRecentEmail = syncItem.Email;
                    mostRecentEmailServerId = syncItem.ServerId;
                }
            }

            // Set the flag status of the most recent email to 2.
            SyncStore syncStore = this.SyncChange(this.LatestSyncKey, new Collection<string>() { mostRecentEmailServerId }, User1Information.InboxCollectionId, null, "2");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R209");

            // Verify MS-ASCON requirement: MS-ASCON_R209
            Site.CaptureRequirementIfIsNotNull(
                syncStore,
                209,
                @"[In Flagging a Conversation for Follow-up] The server sends a Sync command response, as specified in [MS-ASCMD] section 2.2.1.21.");

            // Check whether flag status property is changed to 2 for the most recent flag in the conversation.
            syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            bool setFlags = false;
            foreach (Sync item in syncItems.AddElements)
            {
                if (item.Email.ConversationId == mostRecentEmail.ConversationId && item.Email.Flag.Status == "2")
                {
                    if (item.ServerId == mostRecentEmailServerId)
                    {
                        setFlags = true;
                        continue;
                    }
                    else
                    {
                        setFlags = false;
                        break;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R324");

            // Verify MS-ASCON requirement: MS-ASCON_R324
            Site.CaptureRequirementIfIsTrue(
                setFlags,
                324,
                @"[In Flagging a Conversation for Follow-up] When a conversation is flagged for follow-up, the most recent e-mail message that is in the conversation and that is in the current folder is flagged.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R206");

            // Verify MS-ASCON requirement: MS-ASCON_R206
            Site.CaptureRequirementIfIsTrue(
                setFlags,
                206,
                @"[In Flagging a Conversation for Follow-up] When the server receives a request to flag a conversation for follow-up, as specified in section 3.1.4.2, the server flags the most recent e-mail message that is in the conversation and that is in the current folder.");
            #endregion

            #region Clear flag of the conversation.
            // Set conversation flag's status to 0 for clearing the active flag.
            this.SyncChange(this.LatestSyncKey, conversationItem.ServerId, User1Information.InboxCollectionId, null, "0");

            // Check whether flag of all messages are cleared.
            syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            bool clearFlags = true;
            int itemCount = 0;

            foreach (Sync item in syncItems.AddElements)
            {
                if (conversationItem.ServerId.Contains(item.ServerId))
                {
                    itemCount++;
                    if (!string.IsNullOrEmpty(item.Email.Flag.Status))
                    {
                        clearFlags = false;
                        break;
                    }
                }
            }

            Site.Assert.AreEqual<int>(2, itemCount, "The Sync response should have 2 emails belong to the conversation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R207");

            // Verify MS-ASCON requirement: MS-ASCON_R207
            Site.CaptureRequirementIfIsTrue(
                clearFlags,
                207,
                @"[In Flagging a Conversation for Follow-up] If a flag is cleared on a conversation, the server clears flags on all e-mail messages that are in the conversation and that are in the current folder");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R325");

            // Verify MS-ASCON requirement: MS-ASCON_R325
            Site.CaptureRequirementIfIsTrue(
                clearFlags,
                325,
                @"[In Flagging a Conversation for Follow-up] Clearing a flag on a conversation will clear flags on all e-mail messages that are in the conversation and that are in the current folder.");
            #endregion

            #region Mark complete flag to the conversation.
            // Set conversation flags status to 1 for setting a complete flag.
            this.SyncChange(this.LatestSyncKey, conversationItem.ServerId, User1Information.InboxCollectionId, null, "1");

            // Check whether flag status property is changed to 1 for all messages in the conversation.
            syncItems = this.CallSyncCommand(User1Information.InboxCollectionId, false);

            bool markCompleteFlags = true;
            itemCount = 0;

            foreach (Sync item in syncItems.AddElements)
            {
                if (conversationItem.ServerId.Contains(item.ServerId))
                {
                    itemCount++;
                    if (item.Email.Flag.Status != "1")
                    {
                        markCompleteFlags = false;
                        break;
                    }
                }
            }

            Site.Assert.AreEqual<int>(2, itemCount, "The Sync response should have 2 emails belong to the conversation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R162");

            // Verify MS-ASCON requirement: MS-ASCON_R162
            // If all emails in the conversation are marked as complete, then this requirement can be captured
            Site.CaptureRequirementIfIsTrue(
                markCompleteFlags,
                162,
                @"[In Flagging a Conversation for Follow-up] Marking a flagged conversation as complete will mark all flagged e-mail messages that are in the conversation and that are in the current folder as complete.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R208");

            // Verify MS-ASCON requirement: MS-ASCON_R208
            // If all emails in the conversation are marked as complete, then this requirement can be captured
            Site.CaptureRequirementIfIsTrue(
                markCompleteFlags,
                208,
                @"[In Flagging a Conversation for Follow-up] If a flagged conversation is marked as complete, the server marks all flagged e-mail messages that are in the conversation and that are in the current folder as complete.");
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC03_Sync_Delete
        /// <summary>
        /// This test case is designed to validate deleting a conversation by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC03_Sync_Delete()
        {
            #region Create a conversation and sync to get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem inboxConversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Delete the created conversation by Sync command.
            string[] serverIds = new string[inboxConversationItem.ServerId.Count];
            inboxConversationItem.ServerId.CopyTo(serverIds, 0);
            SyncStore syncStore = this.SyncDelete(User1Information.InboxCollectionId, this.LatestSyncKey, serverIds);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.DeletedItemsCollectionId, conversationSubject, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R201");

            // Verify MS-ASCON requirement: MS-ASCON_R201
            // If the response is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncStore,
                201,
                @"[In Deleting a Conversation] The server sends a Sync command response as specified in [MS-ASCMD] section 2.2.1.21.");
            #endregion

            #region Sync messages in the Inbox and Deleted Items folder after conversation deleted.
            Sync deleteResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);

            ConversationItem deleteConversationItem = this.GetConversationItem(User1Information.DeletedItemsCollectionId, inboxConversationItem.ConversationId);

            // Verify that server moves all e-mail messages that are in the conversation from the Inbox folder to the Deleted Items folder
            bool isDeleted = deleteConversationItem.ServerId.Count == inboxConversationItem.ServerId.Count && deleteResult == null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R199");
            Site.Log.Add(LogEntryKind.Debug, "The count of emails in the specified conversation moved to Deleted Items folder is {0}.", deleteConversationItem.ServerId.Count);
            Site.Log.Add(LogEntryKind.Debug, "The count of emails in the specified conversation deleted from Inbox folder is {0}.", inboxConversationItem.ServerId.Count);

            // Verify MS-ASCON requirement: MS-ASCON_R199
            // If all emails in the conversation are deleted, then this requirement is captured.
            Site.CaptureRequirementIfIsTrue(
                isDeleted,
                199,
                @"[In Deleting a Conversation] When the server receives a request to delete a conversation, as specified in section 3.1.4.1, the server moves all e-mail messages that are in the conversation from the current folder to the Deleted Items folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R157");
            Site.Log.Add(LogEntryKind.Debug, "The count of emails in the specified conversation moved to Deleted Items folder is {0}.", deleteConversationItem.ServerId.Count);
            Site.Log.Add(LogEntryKind.Debug, "The count of emails in the specified conversation deleted from Inbox folder is {0}.", inboxConversationItem.ServerId.Count);

            // Verify MS-ASCON requirement: MS-ASCON_R157
            // If all emails in the conversation are deleted, then this requirement is captured.
            Site.CaptureRequirementIfIsTrue(
                isDeleted,
                157,
                @"[In Deleting a Conversation] When a conversation is deleted, all e-mail messages that are in the conversation are moved from the current folder to the Deleted Items folder.");
            #endregion

            #region User2 replies a message to create a future e-mail message for that conversation.
            // Switch the current user to User2.
            this.SwitchUser(this.User2Information, false);
            Sync syncResult = this.SyncEmail(conversationSubject, User2Information.InboxCollectionId, true, null, null);

            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);

            // Smart reply the received email from User2 to User1.
            this.CallSmartReplyCommand(syncResult.ServerId, User2Information.InboxCollectionId, user2MailboxAddress, user1MailboxAddress, conversationSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, false);

            // Switch the current user to User1.
            this.SwitchUser(this.User1Information, false);
            syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, true, null, null);
            ConversationItem futureInboxConversation = this.GetConversationItem(User1Information.InboxCollectionId, syncResult.Email.ConversationId);

            bool isFutureEmailMoved = futureInboxConversation.ServerId.Count == 1 && futureInboxConversation.ConversationId == inboxConversationItem.ConversationId;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R158");
            Site.Log.Add(LogEntryKind.Debug, "The count of the received emails of the specified conversation in Inbox folder is {0}.", futureInboxConversation.ServerId.Count);
            Site.Log.Add(LogEntryKind.Debug, "The ConversationId of the received email in Inbox folder should be {0}, while it is {1} for the deleted emails.", futureInboxConversation.ConversationId, inboxConversationItem.ConversationId);

            // Verify MS-ASCON requirement: MS-ASCON_R158
            // New email of the conversation is received and it is not moved to Deleted Items folder, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isFutureEmailMoved,
                158,
                @"[In Deleting a Conversation] Future e-mail messages for the same conversation are not affected.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R200");
            Site.Log.Add(LogEntryKind.Debug, "The count of the received emails of the specified conversation in Inbox folder is {0}.", futureInboxConversation.ServerId.Count);
            Site.Log.Add(LogEntryKind.Debug, "The ConversationId of the received email in Inbox folder should be {0}, while it is {1} for the deleted emails.", futureInboxConversation.ConversationId, inboxConversationItem.ConversationId);

            // Verify MS-ASCON requirement: MS-ASCON_R200
            // New email of the conversation is received and it is not moved to Deleted Items folder, so this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isFutureEmailMoved,
                200,
                @"[In Deleting a Conversation] The server does not move future e-mail messages for the conversation.");
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC04_Sync_Filter
        /// <summary>
        /// This test case is designed to validate filtering a conversation by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC04_Sync_Filter()
        {
            #region Create a conversation and sync to get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem inboxConversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call Sync command with setting ConversationMode element to true.
            SyncStore syncStore = this.CallSyncCommand(User1Information.InboxCollectionId, true);

            int itemCount = 0;
            foreach (Sync item in syncStore.AddElements)
            {
                if (item.Email.Subject.Contains(conversationSubject))
                {
                    itemCount++;
                }
            }
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R180");

            // Verify MS-ASCON requirement: MS-ASCON_R180
            // If the count of the items got from Sync command is equal to the count of the item in the conversation, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                inboxConversationItem.ServerId.Count,
                itemCount,
                180,
                @"[In Synchronizing a Conversation] When a conversation is synchronized, all e-mail messages that are part of the conversation and that are in the specified folder are synchronized.");
        }
        #endregion

        #region MSASCON_S01_TC05_Sync_ConversationIndex
        /// <summary>
        /// This test case is designed to validate ConversationIndex element by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC05_Sync_ConversationIndex()
        {
            #region User2 sends mail to User1
            this.SwitchUser(this.User2Information, true);
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            string user1MailboxAddress = Common.GetMailAddress(User1Information.UserName, User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(User2Information.UserName, User2Information.UserDomain);
            this.CallSendMailCommand(user2MailboxAddress, user1MailboxAddress, conversationSubject, null);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, false);
            #endregion

            #region User1 replies to User2
            this.SwitchUser(this.User1Information, false);
            Sync syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, true, null, null);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R304");

            // Verify MS-ASCON requirement: MS-ASCON_R304
            // If the ConversationIndex element is not null, then this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncResult.Email.ConversationIndex,
                304,
                @"[In ConversationIndex] The value of the first timestamp is derived from the date and time when the message was originally sent by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R95");

            // Verify MS-ASCON requirement: MS-ASCON_R95
            // If the ConversationIndex element is not null, then this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                syncResult.Email.ConversationIndex,
                95,
                @"[In Conversation Index Header] The Conversation Index Header value is derived from the date and time when the message was originally sent by the server.");

            byte[] conversationIndexHeader = Convert.FromBase64String(syncResult.Email.ConversationIndex);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R90");

            // Verify MS-ASCON requirement: MS-ASCON_R90
            // If the ConversationIndex element is 5 bytes in length, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                conversationIndexHeader.Length,
                90,
                @"[In ConversationIndex] Conversation Index Header (5 bytes): A Conversation Index Header that is derived from the date and time when the message was originally sent by the server, as specified in section 2.2.2.4.1.");

            bool conversationIdAndIndex = syncResult.Email.ConversationIndex != null && syncResult.Email.ConversationId != null;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R192");

            // Verify MS-ASCON requirement: MS-ASCON_R192
            // If the ConversationIndex element and ConversationId element both exist, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                conversationIdAndIndex,
                192,
                @"[In Abstract Data Model] The server creates a conversation ID and a conversation index on the e-mail item when the user sends an e-mail message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R194");

            // Verify MS-ASCON requirement: MS-ASCON_R194
            // If the ConversationIndex element and ConversationId element both exist, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                conversationIdAndIndex,
                194,
                @"[In Higher-Layer Triggered Events] The server creates a conversation ID and a conversation index on the e-mail item when the user sends an e-mail message.");

            this.CallSmartReplyCommand(syncResult.ServerId, User1Information.InboxCollectionId, user1MailboxAddress, user2MailboxAddress, conversationSubject);
            TestSuiteBase.RecordCaseRelativeItems(this.User2Information, User2Information.InboxCollectionId, conversationSubject, false);
            #endregion

            #region User2 forwards to User3
            this.SwitchUser(this.User2Information, false);
            syncResult = this.SyncEmail(conversationSubject, User2Information.InboxCollectionId, true, null, null);
            byte[] conversationIndexReply = Convert.FromBase64String(syncResult.Email.ConversationIndex);

            bool additionalTimestampAdded = conversationIndexReply.Length > conversationIndexHeader.Length && BitConverter.ToString(conversationIndexReply).Replace("-", string.Empty).StartsWith(BitConverter.ToString(conversationIndexHeader).Replace("-", string.Empty), StringComparison.OrdinalIgnoreCase);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R86");
            Site.Log.Add(LogEntryKind.Debug, "After User1 reply the email to User2, the ConversationIndex is {0} and its length is {1}.", conversationIndexReply, conversationIndexReply.Length);

            // Verify MS-ASCON requirement: MS-ASCON_R86
            // If the ConversationIndex element is longer than ConversationIndexHeader which means additional timestamp has been added to ConversationIndex, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                additionalTimestampAdded,
                86,
                @"[In ConversationIndex] Each additional timestamp specifies the difference between the current time and the time specified by the first timestamp.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R87");
            Site.Log.Add(LogEntryKind.Debug, "After User1 reply the email to User2, the ConversationIndex is {0} and its length is {1}.", conversationIndexReply, conversationIndexReply.Length);

            // Verify MS-ASCON requirement: MS-ASCON_R87
            // If the ConversationIndex element is longer than ConversationIndexHeader which means additional timestamp has been added to ConversationIndex, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                additionalTimestampAdded,
                87,
                @"[In ConversationIndex] Additional timestamps are added when the message is  [forwarded or] replied to.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R91");

            // Verify MS-ASCON requirement: MS-ASCON_R91
            // If the ConversationIndex element is 5 bytes longer than ConversationIndexHeader, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexHeader.Length + 5,
                conversationIndexReply.Length,
                91,
                @"[In ConversationIndex] Response Level 1 (5 bytes): A Response Level that contains information about the time the message was [forwarded or] replied to.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R306");

            // Verify MS-ASCON requirement: MS-ASCON_R306
            // If the ConversationIndex element is 5 bytes longer than ConversationIndexHeader, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexHeader.Length + 5,
                conversationIndexReply.Length,
                306,
                @"[In ConversationIndex] [Response Level 1 (5 bytes):] Additional Response Level fields are added to the email2:ConversationIndex each time the message is [forwarded or] replied to.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R307");

            // Verify MS-ASCON requirement: MS-ASCON_R307
            // If the ConversationIndex element is 5 bytes longer than ConversationIndexHeader, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexHeader.Length + 5,
                conversationIndexReply.Length,
                307,
                @"[In ConversationIndex] Response Level N (5 bytes): Additional Response Level fields for each time the message is [forwarded or] replied to. ");

            string user3MailboxAddress = Common.GetMailAddress(User3Information.UserName, User3Information.UserDomain);

            this.CallSmartForwardCommand(syncResult.ServerId, User2Information.InboxCollectionId, user2MailboxAddress,  user3MailboxAddress, conversationSubject);
            #endregion

            #region User3 gets the email.
            this.SwitchUser(this.User3Information, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User3Information, User3Information.InboxCollectionId, conversationSubject, false);

            syncResult = this.SyncEmail(conversationSubject, User3Information.InboxCollectionId, true, null, null);
            byte[] conversationIndexForward = Convert.FromBase64String(syncResult.Email.ConversationIndex);

            additionalTimestampAdded = conversationIndexForward.Length > conversationIndexReply.Length && BitConverter.ToString(conversationIndexForward).Replace("-", string.Empty).StartsWith(BitConverter.ToString(conversationIndexReply).Replace("-", string.Empty), StringComparison.OrdinalIgnoreCase);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R88");
            Site.Log.Add(LogEntryKind.Debug, "After User2 forward the email to User3, the ConversationIndex is {0} and its length is {1}.", conversationIndexForward, conversationIndexForward.Length);

            // Verify MS-ASCON requirement: MS-ASCON_R88
            // If the ConversationIndex element is longer than the ConversationIndex of the most recent email which means additional timestamp has been added to ConversationIndex, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                additionalTimestampAdded,
                88,
                @"[In ConversationIndex] Additional timestamps are added when the message is forwarded [or replied to].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R105");
            Site.Log.Add(LogEntryKind.Debug, "After User2 forward the email to User3, the ConversationIndex is {0} and its length is {1}.", conversationIndexForward, conversationIndexForward.Length);

            // Verify MS-ASCON requirement: MS-ASCON_R105
            // If the ConversationIndex element is longer than the ConversationIndex of the most recent email which means additional timestamp has been added to ConversationIndex, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                additionalTimestampAdded,
                105,
                @"[In Response Level] The Response Level field contains information about the time the message was forwarded [or replied to].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R259");

            // Verify MS-ASCON requirement: MS-ASCON_R259
            // If the ConversationIndex element is 5 bytes longer than the ConversationIndex of the most recent email, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexReply.Length + 5,
                conversationIndexForward.Length,
                259,
                @"[In ConversationIndex] Response Level 1 (5 bytes): A Response Level that contains information about the time the message was forwarded [or replied to].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R305");

            // Verify MS-ASCON requirement: MS-ASCON_R305
            // If the ConversationIndex element is 5 bytes longer than the ConversationIndex of the most recent email, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexReply.Length + 5,
                conversationIndexForward.Length,
                305,
                @"[In ConversationIndex] [Response Level 1 (5 bytes):] Additional Response Level fields are added to the email2:ConversationIndex each time the message is forwarded [or replied to].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R92");

            // Verify MS-ASCON requirement: MS-ASCON_R92
            // If the ConversationIndex element is 5 bytes longer than the ConversationIndex of the most recent email, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                conversationIndexReply.Length + 5,
                conversationIndexForward.Length,
                92,
                @"[In ConversationIndex] Response Level N (5 bytes): Additional Response Level fields for each time the message is forwarded [or replied to].");
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC06_Sync_NoConversationId
        /// <summary>
        /// This test case is designed to validate the ConversationId element is not present by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC06_Sync_NoConversationId()
        {
            #region Initialize sync and get synckey.
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(User1Information.CalendarCollectionId);
            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);
            #endregion

            #region Create a calendar item and sync to get the ConversationId node in the response xml.
            string calendarSubject = Common.GenerateResourceName(Site, "TestCalendar");
            this.SyncAdd(User1Information.CalendarCollectionId, calendarSubject, syncStore.SyncKey);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.CalendarCollectionId, calendarSubject, false);

            // Call initial Sync command.
            syncRequest = Common.CreateInitialSyncRequest(User1Information.CalendarCollectionId);
            syncStore = this.CONAdapter.Sync(syncRequest);

            // Sync calendar folder
            syncRequest = TestSuiteHelper.GetSyncRequest(User1Information.CalendarCollectionId, syncStore.SyncKey, null, null, false);
            this.CONAdapter.Sync(syncRequest);
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            XmlElement lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
            xmlDoc.LoadXml(lastRawResponse.InnerXml);
            System.Xml.XmlNodeList conversationIdNodes = xmlDoc.GetElementsByTagName("ConversationId");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R77");

            // Verify MS-ASCON requirement: MS-ASCON_R77
            // If the ConversionId node does not exist, then this requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                0,
                conversationIdNodes.Count,
                77,
                @"[In ConversationId (Sync)] The email2:ConversationId element is not present if there is no conversation ID associated with the message.");
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC07_Sync_MessagePart
        /// <summary>
        /// This test case is designed to validate requesting the message part by Sync command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC07_Sync_MessagePart()
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

            #region Call Sync command without BodyPreference or BodyPartPreference element.
            this.SwitchUser(this.User1Information, false);

            // Call Sync command without BodyPreference or BodyPartPreference element.
            Sync syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, null, null);
            this.VerifyMessagePartWithoutPreference(syncItem.Email);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R237");

            // Verify MS-ASCON requirement: MS-ASCON_R237
            Site.CaptureRequirementIfIsNull(
                syncItem.Email.BodyPart,
                237,
                @"[In Sending a Message Part] The airsyncbase:BodyPart element is not present in the [Sync command] response if the client did not request the message part, as specified in section 3.1.4.10.");
            #endregion

            #region Call Sync command with only BodyPreference element.
            Request.BodyPreference bodyPreference = new Request.BodyPreference()
            {
                Type = 2,
            };

            syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, null, bodyPreference);
            this.VerifyMessagePartWithBodyPreference(syncItem.Email);
            #endregion

            #region Call Sync command with only BodyPartPreference element.
            Request.BodyPartPreference bodyPartPreference = new Request.BodyPartPreference()
            {
                Type = 2,
            };

            // Get all the email BodyPart data.
            this.SyncEmail(subject, User1Information.InboxCollectionId, true, bodyPartPreference, null);
            XmlElement lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
            string allData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);

            bodyPartPreference = new Request.BodyPartPreference()
            {
                Type = 2,
                TruncationSize = 12,
                TruncationSizeSpecified = true,
            };

            syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, bodyPartPreference, null);
            lastRawResponse = (XmlElement)this.CONAdapter.LastRawResponseXml;
            string truncatedData = TestSuiteHelper.GetDataInnerText(lastRawResponse, "BodyPart", "Data", subject);
            this.VerifyMessagePartWithBodyPartPreference(syncItem.Email, truncatedData, allData, (int)bodyPartPreference.TruncationSize);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R234");

            // Verify MS-ASCON requirement: MS-ASCON_R234
            Site.CaptureRequirementIfIsNotNull(
                syncItem.Email.BodyPart,
                234,
                @"[In Sending a Message Part] If the client Sync command request ([MS-ASCMD] section 2.2.1.21) [, Search command request ([MS-ASCMD] section 2.2.1.16) or ItemOperations command request 9([MS-ASCMD] section 2.2.1.10)] includes the airsyncbase:BodyPartPreference element (section 2.2.2.2), then the server uses the airsyncbase:BodyPart element (section 2.2.2.1) to encapsulate the message part in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R38");

            // A message part and its meta-data are encapsulated by BodyPart element in the Sync response, so this requirement can be captured.
            Site.CaptureRequirement(
                38,
                @"[In BodyPart] The airsyncbase:BodyPart element ([MS-ASAIRS] section 2.2.2.10) encapsulates a message part and its meta-data in a Sync command response ([MS-ASCMD] section 2.2.1.21) [, an ItemOperations command response ([MS-ASCMD] section 2.2.1.10) or a Search command response ([MS-ASCMD] section 2.2.1.16)].");
            #endregion

            #region Calls Sync command with both BodyPreference and BodyPartPreference elements.
            syncItem = this.SyncEmail(subject, User1Information.InboxCollectionId, true, bodyPartPreference, bodyPreference);
            this.VerifyMessagePartWithBothPreference(syncItem.Email);
            #endregion
        }
        #endregion

        #region MSASCON_S01_TC08_Sync_Status164
        /// <summary>
        /// This test case is designed to validate status 164 is returned if a value other than 2 is specified in the Type element of BodyPartPreference element in Sync command request.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S01_TC08_Sync_Status164()
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

            #region Call sync command with BodyPartPreference element and set the Type element to 3
            this.SwitchUser(this.User1Information, false);

            // Check whether the mail has been received.
            this.SyncEmail(subject, User1Information.InboxCollectionId, true, null, null);

            Request.BodyPartPreference bodyPartPreference = new Request.BodyPartPreference()
            {
                Type = 3,
            };

            // Call initial Sync command.
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(User1Information.InboxCollectionId);
            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            syncRequest = TestSuiteHelper.GetSyncRequest(User1Information.InboxCollectionId, syncStore.SyncKey, bodyPartPreference, null, false);
            syncStore = this.CONAdapter.Sync(syncRequest);
            this.VerifyMessagePartStatus164(syncStore.Status);
            #endregion
        }
        #endregion
    }
}