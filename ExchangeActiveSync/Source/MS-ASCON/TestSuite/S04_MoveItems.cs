namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;

    /// <summary>
    /// This scenario is designed to move a conversation from the current folder using MoveItems command.
    /// </summary>
    [TestClass]
    public class S04_MoveItems : TestSuiteBase
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

        #region MSASCON_S04_TC01_MoveItems_Move
        /// <summary>
        /// This test case is designed to validate moving a conversation by MoveItems command.
        /// </summary>
        [TestCategory("MSASCON"), TestMethod()]
        public void MSASCON_S04_TC01_MoveItems_Move()
        {
            #region Create a conversation and get the created conversation item.
            string conversationSubject = Common.GenerateResourceName(Site, "Conversation");
            ConversationItem sourceConversationItem = this.CreateConversation(conversationSubject);
            #endregion

            #region Call MoveItems command to move the conversation from Inbox folder to SentItems folder.
            MoveItemsResponse moveItemsResponse = this.CallMoveItemsCommand(sourceConversationItem.ServerId, User1Information.InboxCollectionId, User1Information.SentItemsCollectionId);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.InboxCollectionId, conversationSubject, true);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, User1Information.SentItemsCollectionId, conversationSubject, false);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R218");

            // Verify MS-ASCON requirement: MS-ASCON_R218
            Site.CaptureRequirementIfIsNotNull(
                moveItemsResponse.ResponseData,
                218,
                @"[In Processing a MoveItems Command] The server sends a MoveItems command response, as specified in [MS-ASCMD] section 2.2.1.12.");
            #endregion

            #region Synchronize emails in the Inbox folder and Sent Items folder after conversation moved.
            // Call Sync command to get the emails of the conversation in Inbox folder.
            DataStructures.Sync syncResult = this.SyncEmail(conversationSubject, User1Information.InboxCollectionId, false, null, null);

            // Get the emails of the conversation in Sent Items folder.
            ConversationItem destinationCoversationItem = this.GetConversationItem(User1Information.SentItemsCollectionId, sourceConversationItem.ConversationId);

            // If the emails of the conversation in Inbox folder could not be found and the emails count of the conversation in Sent Items folder is equal to Inbox folder, then the original emails have been moved. 
            bool allEmailsMoved = syncResult == null && sourceConversationItem.ServerId.Count == destinationCoversationItem.ServerId.Count;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R170");
            Site.Log.Add(LogEntryKind.Debug, "The count of the emails of the specified conversation in Inbox folder is {0}, while it is {1} in Sent Items folder.", sourceConversationItem.ServerId.Count, destinationCoversationItem.ServerId.Count);

            // Verify MS-ASCON requirement: MS-ASCON_R170
            Site.CaptureRequirementIfIsTrue(
                allEmailsMoved,
                170,
                @"[In Moving a Conversation from the Current Folder] When a conversation is moved from the current folder to another folder, all e-mail messages that are in the conversation are moved from the current folder to the destination folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R217");
            Site.Log.Add(LogEntryKind.Debug, "The count of the emails of the specified conversation in Inbox folder is {0}, while it is {1} in Sent Items folder.", sourceConversationItem.ServerId.Count, destinationCoversationItem.ServerId.Count);

            // Verify MS-ASCON requirement: MS-ASCON_R217
            Site.CaptureRequirementIfIsTrue(
                allEmailsMoved,
                217,
                @"[In Processing a MoveItems Command] When the server receives a request to move a conversation, as specified in section 3.1.4.5, the server moves all e-mail messages that are in the conversation from the current folder to the specified destination folder.");
            #endregion
        }
        #endregion
    }
}