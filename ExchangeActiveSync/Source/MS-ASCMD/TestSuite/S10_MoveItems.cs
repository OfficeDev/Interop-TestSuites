//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// This scenario is designed to test the MoveItems command.
    /// </summary>
    [TestClass]
    public class S10_MoveItems : TestSuiteBase
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
        /// This test case is used to verify if the MoveItems command executes successfully, a new serverID should be assigned by the server.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC01_MoveItems_Success()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server, and get the ServerId of sent email item and the SyncKey
            SyncResponse syncResponseInbox = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            string syncKeyInbox = this.LastSyncKey;
            string serverId = TestSuiteBase.FindServerId(syncResponseInbox, "Subject", subject);
            #endregion

            #region Call method Sync to synchronize changes of DeletedItems folder in User1's mailbox between the client and the server, and get the SyncKey
            this.SyncChanges(this.User1Information.DeletedItemsCollectionId);
            string syncKeyDeletedItems = this.LastSyncKey;
            #endregion

            #region Call method MoveItems with the email item's ServerId to move the email item from Inbox folder to DeletedItems folder.
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, this.User1Information.InboxCollectionId, this.User1Information.DeletedItemsCollectionId);
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4217");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4217
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)3,
                moveItemsResponse.ResponseData.Response[0].Status,
                4217,
                @"[In Status(MoveItems)] [When the scope is Global], [the cause of the status value 3 is] Server successfully completed command.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3977");

            Site.CaptureRequirementIfAreEqual<string>(
                serverId,
                moveItemsResponse.ResponseData.Response[0].SrcMsgId,
                3977,
                "[In SrcMsgId] The SrcMsgId element is a required child element of the Response element in MoveItems command responses that specifies the server ID of the item that was moved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R285");

            Site.CaptureRequirementIfAreNotEqual<string>(
                serverId,
                moveItemsResponse.ResponseData.Response[0].DstMsgId,
                285,
                "[In MoveItems] The MoveItems command moves an item or items from one folder on the server to another [folder].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R290");

            Site.CaptureRequirementIfAreNotEqual<string>(
                serverId,
                moveItemsResponse.ResponseData.Response[0].DstMsgId,
                290,
                "[In MoveItems] An item that has been successfully moved to a different folder can be assigned a new server ID by the server.");
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder and DeletedItems folder in User1's mailbox, and record the item changes for clean up
            bool isItemDeleted = this.CheckDeleteInSyncResponse(syncKeyInbox, this.User1Information.InboxCollectionId, serverId);
            Site.Assert.IsTrue(isItemDeleted, "The item with ServerId: {0} should be deleted for collection ID: {0}.", serverId, this.User1Information.InboxCollectionId);
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);

            bool isItemAdded = this.CheckAddInSyncResponse(syncKeyDeletedItems, this.User2Information.DeletedItemsCollectionId, subject);

            Site.Assert.IsTrue(isItemAdded, "The item with ServerId: {0} should be added for Deleted Items folder", serverId);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.DeletedItemsCollectionId, subject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R289");

            // Verify MS-ASCMD requirement: MS-ASCMD_R289
            // This requirement can be captured directly after previous step when synchronize changes of Inbox folder, there is a Delete operation,
            // and synchronize changes of DeletedItems folder, there is an Add operation.
            Site.CaptureRequirement(
                289,
                @"[In MoveItems] When items are moved between folders on the server, the client receives Delete (section 2.2.3.42) and Add (section 2.2.3.7) operations the next time the client synchronizes the affected folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5879");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5879
            // This requirement can be captured directly after previous step when synchronize changes of Inbox folder, there is a Delete operation on the source item
            // and synchronize changes of DeletedItems folder, there is an Add operation on the item with same subject.
            Site.CaptureRequirement(
                5879,
                @"[In SrcFldId] [The SrcFldId element] specifies the server ID of the source folder (that is, the folder that contains the items to be moved).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MoveItems command, if the request contains invalid source collection ID, the status in response is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC02_MoveItems_Status1_InvalidSrcFldId()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server, and get the ServerId of sent email item
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method MoveItems with the email item's ServerId to move the email item from an invalid source folder to DeletedItems folder.
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, "Invalid SrcFldId", this.User1Information.DeletedItemsCollectionId);
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4201");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4201
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)1,
                moveItemsResponse.ResponseData.Response[0].Status,
                4201,
                @"[In Status(MoveItems)] If the command failed, Status contains a code indicating the type of failure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4207");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4207
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)1,
                moveItemsResponse.ResponseData.Response[0].Status,
                4207,
                @"[In Status(MoveItems)] [When the scope is Item], [the cause of the status value 1 is] The source folder collection ID (CollectionId element (section 2.2.3.30.5) value) is not recognized by the server, possibly because the source folder has been deleted.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MoveItems command, if the request contains an invalid source Item ID, the status in response is equal to 1.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC03_MoveItems_Status1_InvalidSrcMsgId()
        {
            #region User2 calls SendMail command to send a mail to User1.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region User1 calls Sync command to synchronize changes of Inbox folder and get the ServerId of sent email item and the latest SyncKey.
            SyncResponse syncResponseInbox = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponseInbox, "Subject", subject);
            #endregion

            #region User1 calls MoveItems command with the received email item's ServerId to move the email item from Inbox folder to DeletedItems folder.
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, this.User1Information.InboxCollectionId, this.User1Information.DeletedItemsCollectionId);
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);
            TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);
            TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.DeletedItemsCollectionId, subject);
            #endregion

            #region User1 calls MoveItems command with the received email item's ServerId again to move the email item from Inbox folder to DeletedItems folder after the email item is not exist in Inbox folder.
            moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, this.User1Information.InboxCollectionId, this.User1Information.DeletedItemsCollectionId);
            moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4208");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4208
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)1,
                moveItemsResponse.ResponseData.Response[0].Status,
                4208,
                @"[In Status(MoveItems)] [When the scope is Item], [the cause of the status value 1 is] Or, the item with the Item ID (SrcMsgId element (section 2.2.3.160)) has been previously moved out of the folder with the Folder ID (SrcFldId element (section 2.2.3.159)).");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify MoveItems command, if the request contains invalid destination collection ID, the status in response is equal to 2.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC04_MoveItems_Status2()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method MoveItems with the email item's ServerId to move the email item from Inbox folder to an invalid destination folder.
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, this.User1Information.InboxCollectionId, "Invalid DstFldId");
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4212");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4212
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)2,
                moveItemsResponse.ResponseData.Response[0].Status,
                4212,
                @"[In Status(MoveItems)] [When the scope is Item], [the cause of the status value 2 is] The destination folder collection ID (CollectionId element value) is not recognized by the server, possibly because the source folder has been deleted.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MoveItems command, if the source and destination collection IDs are the same, the status in response is equal to 4.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC05_MoveItems_Status4()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponse = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string serverId = TestSuiteBase.FindServerId(syncResponse, "Subject", subject);
            #endregion

            #region Call method MoveItems with the email item's ServerId to move the email item from Inbox folder to Inbox folder.
            MoveItemsRequest moveItemsRequest = TestSuiteBase.CreateMoveItemsRequest(serverId, this.User1Information.InboxCollectionId, this.User1Information.InboxCollectionId);
            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 1);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4219");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4219
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)4,
                moveItemsResponse.ResponseData.Response[0].Status,
                4219,
                @"[In Status(MoveItems)] [When the scope is Item], [the cause of the status value 4 is] The client supplied a destination folder that is the same as the source.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the MoveItems command, if the request contains more than one destination collection ID, the status in response is equal to 5.
        /// </summary>
        [TestCategory("MSASCMD"), TestMethod()]
        public void MSASCMD_S10_TC06_MoveItems_Status5()
        {
            #region User2 sends mail to User1 and does FolderSync in User1's mailbox.
            string subject = this.SendMailAndFolderSync();
            #endregion

            #region Call method Sync to synchronize changes of Inbox folder in User1's mailbox between the client and the server.
            SyncResponse syncResponseInbox = this.GetMailItem(this.User1Information.InboxCollectionId, subject);
            string syncKeyInbox = this.LastSyncKey;
            string serverId = TestSuiteBase.FindServerId(syncResponseInbox, "Subject", subject);
            #endregion

            #region Call method Sync to synchronize changes of DeletedItems folder in User1's mailbox between the client and the server.
            this.SyncChanges(this.User1Information.DeletedItemsCollectionId);
            string syncKeyDeletedItems = this.LastSyncKey;
            #endregion

            #region Call method Sync to synchronize changes of SentItems folder in User1's mailbox between the client and the server.
            this.SyncChanges(this.User1Information.SentItemsCollectionId);
            string syncKeySentItems = this.LastSyncKey;
            #endregion

            #region Call method MoveItems with the email item's ServerId and two different destination collection IDs to move the same email item from Inbox folder to two different folders, and check if there is one response with Status element equal to 5
            MoveItemsRequest moveItemsRequest = new MoveItemsRequest();
            Request.MoveItems moveItems = new Request.MoveItems();
            List<Request.MoveItemsMove> moveItemList = new List<Request.MoveItemsMove>();

            Request.MoveItemsMove moveItemsMoveOne = new Request.MoveItemsMove
            {
                DstFldId = this.User1Information.DeletedItemsCollectionId,
                SrcFldId = this.User1Information.InboxCollectionId,
                SrcMsgId = serverId
            };
            moveItemList.Add(moveItemsMoveOne);

            Request.MoveItemsMove moveItemsMoveTwo = new Request.MoveItemsMove
            {
                DstFldId = this.User1Information.SentItemsCollectionId,
                SrcFldId = this.User1Information.InboxCollectionId,
                SrcMsgId = serverId
            };

            moveItemList.Add(moveItemsMoveTwo);

            moveItems.Move = moveItemList.ToArray();

            moveItemsRequest.RequestData = moveItems;

            MoveItemsResponse moveItemsResponse = this.CMDAdapter.MoveItems(moveItemsRequest);

            this.CheckMoveItemsResponse(moveItemsResponse, 2);

            bool hasStatus5 = false;
            bool hasStatus3 = false;
            foreach (Response.MoveItemsResponse response in moveItemsResponse.ResponseData.Response)
            {
                if (response.Status == 5)
                {
                    hasStatus5 = true;
                    Site.Log.Add(LogEntryKind.Debug, "There should be at least one Status element equal to 5 in MoveItems response");
                }
                else if (response.Status == 3)
                {
                    hasStatus3 = true;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4222");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4222
            Site.CaptureRequirementIfIsTrue(
                hasStatus5,
                4222,
                @"[In Status(MoveItems)] [When the scope is Global], [the cause of the status value 5 is] More than one DstFldId element (section 2.2.3.47.2) was included in the request [or an item with that name already exists].");

            #endregion

            #region Call method Sync to synchronize changes of Inbox folder, DeletedItems and SentItems folder in User1's mailbox, and record the item changes in case there is another success response.
            if (hasStatus3)
            {
                bool isItemDeleted = this.CheckDeleteInSyncResponse(syncKeyInbox, this.User1Information.InboxCollectionId, serverId);
                Site.Assert.IsTrue(isItemDeleted, "The item with ServerId: {0} should be deleted for collection ID: {1}.", serverId, this.User1Information.InboxCollectionId);
                TestSuiteBase.RemoveRecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject);

                bool isItemAddedDeleteItems = this.CheckAddInSyncResponse(syncKeyDeletedItems, this.User1Information.DeletedItemsCollectionId, subject);
                if (isItemAddedDeleteItems)
                {
                    Site.Log.Add(LogEntryKind.Debug, "The item with ServerId: {0} has been added for Deleted Items folder", serverId);
                    TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.DeletedItemsCollectionId, subject);
                }

                bool isItemAddedSentItems = this.CheckAddInSyncResponse(syncKeySentItems, this.User1Information.SentItemsCollectionId, subject);
                if (isItemAddedSentItems)
                {
                    Site.Log.Add(LogEntryKind.Debug, "The item with ServerId: {0} has been added for Sent Items folder", serverId);
                    TestSuiteBase.RecordCaseRelativeItems(this.User1Information, this.User1Information.SentItemsCollectionId, subject);
                }

                Site.Assert.AreNotEqual<bool>(
                    isItemAddedDeleteItems,
                    isItemAddedSentItems,
                    "If the source item is moved successfully, it should exist in either Deleted Items folder or Sent Items folder, but should not exist in both of these folders.");
            }
            #endregion
        }

        #endregion

        #region Private Method
        /// <summary>
        /// This method is used to check the response of MoveItems command.
        /// </summary>
        /// <param name="moveItemsResponse">The response of MoveItems command.</param>
        /// <param name="expectedResponseNumber">The expected number of Response element in MoveItems response</param>
        private void CheckMoveItemsResponse(MoveItemsResponse moveItemsResponse, int expectedResponseNumber)
        {
            Site.Assert.IsNotNull(moveItemsResponse.ResponseData.Response, "The Response elements in MoveItems Response should be not null.");
            Site.Assert.AreEqual<int>(
                expectedResponseNumber,
                moveItemsResponse.ResponseData.Response.Length,
                "There should be {0} Response element(s) in MoveItems Response.",
                expectedResponseNumber);

            for (int i = 0; i < expectedResponseNumber; i++)
            {
                Site.Assert.IsNotNull(moveItemsResponse.ResponseData.Response[i], "The No.{0} Response element in MoveItems Response should be not null.", i + 1);
            }
        }

        /// <summary>
        /// This method is used to retrieve the Commands element in Sync response.
        /// </summary>
        /// <param name="syncResponse">The response of sync command.</param>
        /// <returns>The commands element in the response of sync command.</returns>
        private Response.SyncCollectionsCollectionCommands GetCommandsFromSyncResponse(SyncResponse syncResponse)
        {
            if (syncResponse.ResponseData.Item == null)
            {
                return null;
            }

            Response.SyncCollections syncCollections = (Response.SyncCollections)syncResponse.ResponseData.Item;

            Site.Assert.AreEqual<int>(1, syncCollections.Collection.Length, "There should be one Collection element in Sync response.");
            Site.Assert.IsNotNull(syncCollections.Collection[0], "The Collection element in Sync response should be not null.");

            for (int i = 0; i < syncCollections.Collection[0].ItemsElementName.Length; i++)
            {
                if (Response.ItemsChoiceType10.Commands == syncCollections.Collection[0].ItemsElementName[i])
                {
                    Site.Assert.IsNotNull(syncCollections.Collection[0].Items[i], "The Commands element in Sync response should be not null.");
                    return (Response.SyncCollectionsCollectionCommands)syncCollections.Collection[0].Items[i];
                }
            }

            return null;
        }

        /// <summary>
        /// This method is used to check the Delete element in Sync response.
        /// </summary>
        /// <param name="syncKey">The sync key</param>
        /// <param name="collectionId">Folder's collectionID</param>
        /// <param name="serverId">The ServerId of item which is expected to delete</param>
        /// <returns>The boolean value which indicates the Delete element is found in Sync response or not</returns>
        private bool CheckDeleteInSyncResponse(string syncKey, string collectionId, string serverId)
        {
            SyncResponse syncResponse = this.SyncChanges(syncKey, collectionId);
            Response.SyncCollectionsCollectionCommands commands = this.GetCommandsFromSyncResponse(syncResponse);

            Site.Assert.IsNotNull(commands.Delete, "The Delete elements in Sync response for collection ID: {0} should not be null.", collectionId);
            foreach (Response.SyncCollectionsCollectionCommandsDelete delete in commands.Delete)
            {
                Site.Assert.IsNotNull(delete, "The Delete element in Sync response for collection ID: {0} should not be null.", collectionId);
                if (serverId.Equals(delete.ServerId))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// This method is used to check the Add element in Sync response.
        /// </summary>
        /// <param name="syncKey">The sync key</param>
        /// <param name="collectionId">Folder's collectionID</param>
        /// <param name="subject">The subject of item which is expected to add</param>
        /// <returns>The boolean value which indicates the Add element is found in Sync response or not</returns>
        private bool CheckAddInSyncResponse(string syncKey, string collectionId, string subject)
        {
            SyncResponse syncResponse = syncKey != null
                ? this.SyncChanges(syncKey, collectionId)
                : this.SyncChanges(collectionId);

            Response.SyncCollectionsCollectionCommands commands = this.GetCommandsFromSyncResponse(syncResponse);

            if (commands == null || commands.Add == null)
            {
                return false;
            }
            else
            {
                Site.Assert.IsNotNull(commands.Add, "The Add element in Sync response for collection ID: {0} should not be null", collectionId);
            }

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                Site.Assert.IsNotNull(add, "The Add element in Sync response for collection ID: {0} should not be null", collectionId);
                for (int i = 0; i < add.ApplicationData.Items.Length; i++)
                {
                    if (add.ApplicationData.ItemsElementName[i] == Response.ItemsChoiceType8.Subject1)
                    {
                        if (subject.Equals(add.ApplicationData.Items[i].ToString()))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
        #endregion
    }
}