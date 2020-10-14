namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets or sets the information of User1.
        /// </summary>
        protected UserInformation User1Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User2.
        /// </summary>
        protected UserInformation User2Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User3.
        /// </summary>
        protected UserInformation User3Information { get; set; }

        /// <summary>
        /// Gets MS-ASCON protocol adapter.
        /// </summary>
        protected IMS_ASCONAdapter CONAdapter { get; private set; }

        /// <summary>
        /// Gets the latest SyncKey.
        /// </summary>
        protected string LatestSyncKey { get; private set; }
        #endregion

        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userInformation">The information of the user.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impacts.</param>
        /// <param name="itemSubject">The subject of items that the current test case impacts.</param>
        /// <param name="isDeleted">Whether the item has been deleted and should be removed from the record.</param>
        protected static void RecordCaseRelativeItems(UserInformation userInformation, string folderCollectionId, string itemSubject, bool isDeleted)
        {
            // Record the item in the specified folder.
            CreatedItems items = new CreatedItems { CollectionId = folderCollectionId };
            items.ItemSubject.Add(itemSubject);
            bool isSame = false;

            if (!isDeleted)
            {
                if (userInformation.UserCreatedItems.Count > 0)
                {
                    foreach (CreatedItems createdItems in userInformation.UserCreatedItems)
                    {
                        if (createdItems.CollectionId == folderCollectionId && createdItems.ItemSubject[0] == itemSubject)
                        {
                            isSame = true;
                        }
                    }

                    if (!isSame)
                    {
                        userInformation.UserCreatedItems.Add(items);
                    }
                }
                else
                {
                    userInformation.UserCreatedItems.Add(items);
                }
            }
            else
            {
                if (userInformation.UserCreatedItems.Count > 0)
                {
                    foreach (CreatedItems existItem in userInformation.UserCreatedItems)
                    {
                        if (existItem.CollectionId == folderCollectionId && existItem.ItemSubject[0] == itemSubject)
                        {
                            userInformation.UserCreatedItems.Remove(existItem);
                            break;
                        }
                    }
                }
            }
        }

        #region Test case initialize and cleanup
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.CONAdapter = Site.GetAdapter<IMS_ASCONAdapter>();

            // If implementation doesn't support this specification [MS-ASCON], the case will not start.
            if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-ASCON_Supported", this.Site)))
            {
                this.Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-ASCON_Supported value is set to false in MS-ASCON_{0}_SHOULDMAY.deployment.ptfconfig file.", Common.GetSutVersion(this.Site));
            }

            Site.Assume.AreNotEqual<string>("12.1", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The airsyncbase:BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 12.1. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");

            this.User1Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            this.User2Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            this.User3Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User3Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User3Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                // Switch the current user to User1 and synchronize the folder hierarchy of User1.
                this.SwitchUser(this.User1Information, true);
            }
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // If implementation doesn't support this specification [MS-ASCON], the case will not start.
            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASCON_Supported", this.Site)))
            {
                if (this.User1Information.UserCreatedItems.Count != 0)
                {
                    // Switch to User1
                    this.SwitchUser(this.User1Information, false);
                    this.DeleteItemsInFolder(this.User1Information.UserCreatedItems);
                }

                if (this.User2Information.UserCreatedItems.Count != 0)
                {
                    // Switch to User2
                    this.SwitchUser(this.User2Information, false);
                    this.DeleteItemsInFolder(this.User2Information.UserCreatedItems);
                }

                if (this.User3Information.UserCreatedItems.Count != 0)
                {
                    // Switch to User3
                    this.SwitchUser(this.User3Information, false);
                    this.DeleteItemsInFolder(this.User3Information.UserCreatedItems);
                }
            }

            base.TestCleanup();
        }
        #endregion

        #region Test case base methods

        /// <summary>
        /// Change user to call active sync operations and resynchronize the collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of the user.</param>
        /// <param name="isFolderSyncNeeded">A boolean value that indicates whether needs to synchronize the folder hierarchy.</param>
        protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
        {
            this.CONAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (isFolderSyncNeeded)
            {
                // Call FolderSync command to synchronize the collection hierarchy.
                FolderSyncRequest folderSyncRequest = Common.CreateFolderSyncRequest("0");
                FolderSyncResponse folderSyncResponse = this.CONAdapter.FolderSync(folderSyncRequest);

                // Verify FolderSync command response.
                Site.Assert.AreEqual<int>(
                    1,
                    int.Parse(folderSyncResponse.ResponseData.Status),
                    "If the FolderSync command executes successfully, the Status in response should be 1.");

                // Get the folder collectionId of User1
                if (userInformation.UserName == this.User1Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User1Information.InboxCollectionId))
                    {
                        this.User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.DeletedItemsCollectionId))
                    {
                        this.User1Information.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.CalendarCollectionId))
                    {
                        this.User1Information.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.SentItemsCollectionId))
                    {
                        this.User1Information.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.RecipientInformationCacheCollectionId))
                    {
                        this.User1Information.RecipientInformationCacheCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.RecipientInformationCache, this.Site);
                    }
                }

                // Get the folder collectionId of User2
                if (userInformation.UserName == this.User2Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User2Information.InboxCollectionId))
                    {
                        this.User2Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }
                }

                // Get the folder collectionId of User3
                if (userInformation.UserName == this.User3Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User3Information.InboxCollectionId))
                    {
                        this.User3Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }
                }
            }
        }

        /// <summary>
        /// Create a conversation.
        /// </summary>
        /// <param name="subject">The subject of the emails in the conversation.</param>
        /// <returns>The created conversation item.</returns>
        protected ConversationItem CreateConversation(string subject)
        {
            #region Send email from User2 to User1
            this.SwitchUser(this.User2Information, true);
            string user1MailboxAddress = Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain);
            string user2MailboxAddress = Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain);
            this.CallSendMailCommand(user2MailboxAddress, user1MailboxAddress, subject, null);
            RecordCaseRelativeItems(this.User1Information, this.User1Information.InboxCollectionId, subject, false);
            #endregion

            #region SmartReply the received email from User1 to User2.
            this.SwitchUser(this.User1Information, false);
            Sync syncResult = this.SyncEmail(subject, this.User1Information.InboxCollectionId, true, null, null);

            this.CallSmartReplyCommand(syncResult.ServerId, this.User1Information.InboxCollectionId, user1MailboxAddress, user2MailboxAddress, subject);
            RecordCaseRelativeItems(this.User2Information, this.User2Information.InboxCollectionId, subject, false);
            #endregion

            #region SmartReply the received email from User2 to User1.
            this.SwitchUser(this.User2Information, false);
            syncResult = this.SyncEmail(subject, this.User2Information.InboxCollectionId, true, null, null);
            this.CallSmartReplyCommand(syncResult.ServerId, this.User2Information.InboxCollectionId, user2MailboxAddress, user1MailboxAddress, subject);
            #endregion

            #region Switch current user to User1 and get the conversation item.
            this.SwitchUser(this.User1Information, false);

            int counter = 0;
            int itemsCount;
            int retryLimit = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            do
            {
                System.Threading.Thread.Sleep(waitTime);
                SyncStore syncStore = this.CallSyncCommand(this.User1Information.InboxCollectionId, false);

                // Reset the item count.
                itemsCount = 0;

                foreach (Sync item in syncStore.AddElements)
                {
                    if (item.Email.Subject.Contains(subject))
                    {
                        syncResult = item;
                        itemsCount++;
                    }
                }

                counter++;
            }
            while (itemsCount < 2 && counter < retryLimit);

            Site.Assert.AreEqual<int>(2, itemsCount, "There should be 2 emails with subject {0} in the Inbox folder, actual {1}.", subject, itemsCount);

            return this.GetConversationItem(this.User1Information.InboxCollectionId, syncResult.Email.ConversationId);
            #endregion
        }

        /// <summary>
        /// Call Search command to find a specified conversation.
        /// </summary>
        /// <param name="conversationId">The ConversationId of the items to search.</param>
        /// <param name="itemsCount">The count of the items expected to be found.</param>
        /// <param name="bodyPartPreference">The BodyPartPreference element.</param>
        /// <param name="bodyPreference">The BodyPreference element.</param>
        /// <returns>The SearchStore instance that contains the search result.</returns>
        protected SearchStore CallSearchCommand(string conversationId, int itemsCount, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
        {
            // Create Search command request.
            SearchRequest searchRequest = TestSuiteHelper.GetSearchRequest(conversationId, bodyPartPreference, bodyPreference);
            SearchStore searchStore = this.CONAdapter.Search(searchRequest, true, itemsCount);

            Site.Assert.AreEqual("1", searchStore.Status, "The Search operation should be success.");

            return searchStore;
        }

        /// <summary>
        /// Call SendMail command to send mail.
        /// </summary>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="to">The mailbox address of recipient.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body content of the email.</param>
        protected void CallSendMailCommand(string from, string to, string subject, string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                body = Common.GenerateResourceName(this.Site, "body");
            }

            // Create the SendMail command request.
            string template =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/html; charset=""us-ascii""
MIME-Version: 1.0

<html>
<body>
<font color=""blue"">{3}</font>
</body>
</html>
";

            string mime = Common.FormatString(template, from, to, subject, body);
            SendMailRequest sendMailRequest = Common.CreateSendMailRequest(null, System.Guid.NewGuid().ToString(), mime);

            // Call SendMail command.
            SendMailResponse sendMailResponse = this.CONAdapter.SendMail(sendMailRequest);

            Site.Assert.AreEqual(string.Empty, sendMailResponse.ResponseDataXML, "The SendMail command should be executed successfully.");
        }

        /// <summary>
        /// Call SmartReply command to reply an email.
        /// </summary>
        /// <param name="itemServerId">The ServerId of the email to reply.</param>
        /// <param name="collectionId">The folder collectionId of the source email.</param>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="replyTo">The mailbox address of recipient.</param>
        /// <param name="subject">The subject of the email to reply.</param>
        protected void CallSmartReplyCommand(string itemServerId, string collectionId, string from, string replyTo, string subject)
        {
            // Create SmartReply command request.
            Request.Source source = new Request.Source();
            string mime = Common.CreatePlainTextMime(from, replyTo, string.Empty, string.Empty, subject, "SmartReply content");
            SmartReplyRequest smartReplyRequest = Common.CreateSmartReplyRequest(null, System.Guid.NewGuid().ToString(), mime, source);

            // Set the command parameters.
            smartReplyRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>());

            source.FolderId = collectionId;
            source.ItemId = itemServerId;
            smartReplyRequest.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
            smartReplyRequest.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
            smartReplyRequest.RequestData.ReplaceMime = string.Empty;

            // Call SmartReply command.
            SmartReplyResponse smartReplyResponse = this.CONAdapter.SmartReply(smartReplyRequest);

            Site.Assert.AreEqual(string.Empty, smartReplyResponse.ResponseDataXML, "The SmartReply command should be executed successfully.");
        }

        /// <summary>
        /// Call SmartForward command to forward an email.
        /// </summary>
        /// <param name="itemServerId">The ServerId of the email to reply.</param>
        /// <param name="collectionId">The folder collectionId of the source email.</param>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="forwardTo">The mailbox address of recipient.</param>
        /// <param name="subject">The subject of the email to reply.</param>
        protected void CallSmartForwardCommand(string itemServerId, string collectionId, string from, string forwardTo, string subject)
        {
            // Create SmartForward command request.
            Request.Source source = new Request.Source();
            string mime = Common.CreatePlainTextMime(from, forwardTo, string.Empty, string.Empty, subject, "SmartForward content");
            SmartForwardRequest smartForwardRequest = Common.CreateSmartForwardRequest(null, System.Guid.NewGuid().ToString(), mime, source);

            // Set the command parameters.
            smartForwardRequest.SetCommandParameters(new Dictionary<CmdParameterName, object>());

            source.FolderId = collectionId;
            source.ItemId = itemServerId;
            smartForwardRequest.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
            smartForwardRequest.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
            smartForwardRequest.RequestData.ReplaceMime = string.Empty;

            // Call SmartForward command.
            SmartForwardResponse smartForwardResponse = this.CONAdapter.SmartForward(smartForwardRequest);

            Site.Assert.AreEqual(string.Empty, smartForwardResponse.ResponseDataXML, "The SmartForward command should be executed successfully.");
        }

        /// <summary>
        /// Call ItemOperations command to fetch an email in the specific folder.
        /// </summary>
        /// <param name="collectionId">The folder collection id to be fetched.</param>
        /// <param name="serverId">The ServerId of the item</param>
        /// <param name="bodyPartPreference">The BodyPartPreference element.</param>
        /// <param name="bodyPreference">The bodyPreference element.</param>
        /// <returns>An Email instance that includes the fetch result.</returns>
        protected Email ItemOperationsFetch(string collectionId, string serverId, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
        {
            ItemOperationsRequest itemOperationsRequest = TestSuiteHelper.GetItemOperationsRequest(collectionId, serverId, bodyPartPreference, bodyPreference);
            ItemOperationsResponse itemOperationsResponse = this.CONAdapter.ItemOperations(itemOperationsRequest);
            Site.Assert.AreEqual("1", itemOperationsResponse.ResponseData.Status, "The ItemOperations operation should be success.");

            ItemOperationsStore itemOperationsStore = Common.LoadItemOperationsResponse(itemOperationsResponse);
            Site.Assert.AreEqual(1, itemOperationsStore.Items.Count, "Only one email is supposed to be fetched.");
            Site.Assert.AreEqual("1", itemOperationsStore.Items[0].Status, "The fetch result should be success.");
            Site.Assert.IsNotNull(itemOperationsStore.Items[0].Email, "The fetched email should not be null.");

            return itemOperationsStore.Items[0].Email;
        }

        /// <summary>
        /// Call ItemOperations command to move a conversation to a folder.
        /// </summary>
        /// <param name="conversationId">The Id of conversation to be moved.</param>
        /// <param name="destinationFolder">The destination folder id.</param>
        /// <param name="moveAlways">Should future messages always be moved.</param>
        /// <returns>An instance of the ItemOperationsResponse.</returns>
        protected ItemOperationsResponse ItemOperationsMove(string conversationId, string destinationFolder, bool moveAlways)
        {
            Request.ItemOperationsMove move = new Request.ItemOperationsMove
            {
                DstFldId = destinationFolder,
                ConversationId = conversationId
            };

            if (moveAlways)
            {
                move.Options = new Request.ItemOperationsMoveOptions { MoveAlways = string.Empty };
            }

            ItemOperationsRequest itemOperationRequest = Common.CreateItemOperationsRequest(new object[] { move });
            ItemOperationsResponse itemOperationResponse = this.CONAdapter.ItemOperations(itemOperationRequest);

            Site.Assert.AreEqual("1", itemOperationResponse.ResponseData.Status, "The ItemOperations operation should be success.");
            Site.Assert.AreEqual(1, itemOperationResponse.ResponseData.Response.Move.Length, "The server should return one Move element in ItemOperationsResponse.");

            return itemOperationResponse;
        }

        /// <summary>
        /// Call Sync command to add items to the specified folder.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the specified folder.</param>
        /// <param name="subject">Subject of the item to add.</param>
        /// <param name="syncKey">The latest SyncKey.</param>
        protected void SyncAdd(string collectionId, string subject, string syncKey)
        {
            // Create Sync request.
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = Guid.NewGuid().ToString(),
                ApplicationData = new Request.SyncCollectionAddApplicationData
                {
                    Items = new object[] { subject },
                    ItemsElementName = new Request.ItemsChoiceType8[] { Request.ItemsChoiceType8.Subject2 }
                }
            };

            Request.SyncCollection collection = new Request.SyncCollection
            {
                Commands = new object[] { add },
                CollectionId = collectionId,
                SyncKey = syncKey
            };
            SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });

            // Call Sync command to add the item.
            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            // Verify Sync command response.
            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            this.LatestSyncKey = syncStore.SyncKey;
        }

        /// <summary>
        /// Call Sync command to change the status of the emails in Inbox folder.
        /// </summary>
        /// <param name="syncKey">The latest SyncKey.</param>
        /// <param name="serverIds">The collection of ServerIds.</param>
        /// <param name="collectionId">The folder collectionId which needs to be sychronized.</param>
        /// <param name="read">Read element of the item.</param> 
        /// <param name="status">Flag status of the item.</param> 
        /// <returns>The SyncStore instance returned from Sync command.</returns>
        protected SyncStore SyncChange(string syncKey, Collection<string> serverIds, string collectionId, bool? read, string status)
        {
            List<Request.SyncCollectionChange> changes = new List<Request.SyncCollectionChange>();

            foreach (string serverId in serverIds)
            {
                Request.SyncCollectionChange change = new Request.SyncCollectionChange
                {
                    ServerId = serverId,
                    ApplicationData = new Request.SyncCollectionChangeApplicationData()
                };

                List<object> changeItems = new List<object>();
                List<Request.ItemsChoiceType7> changeItemsElementName = new List<Request.ItemsChoiceType7>();

                if (read != null)
                {
                    changeItems.Add(read);
                    changeItemsElementName.Add(Request.ItemsChoiceType7.Read);
                }

                if (!string.IsNullOrEmpty(status))
                {
                    Request.Flag flag = new Request.Flag();
                    if (status == "1")
                    {
                        // The Complete Time format is yyyy-MM-ddThh:mm:ss.fffZ.
                        flag.CompleteTime = System.DateTime.Now.ToUniversalTime();
                        flag.CompleteTimeSpecified = true;
                        flag.DateCompleted = System.DateTime.Now.ToUniversalTime();
                        flag.DateCompletedSpecified = true;
                    }

                    flag.Status = status;
                    flag.FlagType = "Flag for follow up";

                    changeItems.Add(flag);
                    changeItemsElementName.Add(Request.ItemsChoiceType7.Flag);
                }

                change.ApplicationData.Items = changeItems.ToArray();
                change.ApplicationData.ItemsElementName = changeItemsElementName.ToArray();

                changes.Add(change);
            }

            Request.SyncCollection collection = new Request.SyncCollection
            {
                CollectionId = collectionId,
                SyncKey = syncKey,
                Commands = changes.ToArray()
            };

            SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });
            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            // Verify Sync command response.
            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            this.LatestSyncKey = syncStore.SyncKey;

            return syncStore;
        }

        /// <summary>
        /// Call Sync command to delete items.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder.</param>
        /// <param name="syncKey">The latest SyncKey.</param> 
        /// <param name="serverIds">The ServerId of the items to delete.</param> 
        /// <returns>The SyncStore instance returned from Sync command.</returns>
        protected SyncStore SyncDelete(string collectionId, string syncKey, string[] serverIds)
        {
            List<Request.SyncCollectionDelete> deleteCollection = new List<Request.SyncCollectionDelete>();
            foreach (string itemId in serverIds)
            {
                Request.SyncCollectionDelete delete = new Request.SyncCollectionDelete { ServerId = itemId };
                deleteCollection.Add(delete);
            }

            Request.SyncCollection collection = new Request.SyncCollection
            {
                Commands = deleteCollection.ToArray(),
                DeletesAsMoves = true,
                DeletesAsMovesSpecified = true,
                CollectionId = collectionId,
                SyncKey = syncKey
            };

            SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { collection });

            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            // Verify Sync command response.
            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncStore;
        }

        /// <summary>
        /// Find the specified email.
        /// </summary>
        /// <param name="subject">The subject of the email to find.</param>
        /// <param name="collectionId">The folder collectionId which needs to be synchronized.</param>
        /// <param name="isRetryNeeded">A Boolean value indicates whether need retry.</param>
        /// <param name="bodyPartPreference">The bodyPartPreference in the options element.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <returns>The found email object.</returns>
        protected Sync SyncEmail(string subject, string collectionId, bool isRetryNeeded, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
        {
            // Call initial Sync command.
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(collectionId);
            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            // Find the specific email.
            syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, bodyPartPreference, bodyPreference, false);
            Sync syncResult = this.CONAdapter.SyncEmail(syncRequest, subject, isRetryNeeded);

            this.LatestSyncKey = syncStore.SyncKey;

            return syncResult;
        }

        /// <summary>
        /// Sync items in the specified folder.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder.</param>
        /// <param name="conversationMode">The value of ConversationMode element.</param>
        /// <returns>A SyncStore instance that contains the result.</returns>
        protected SyncStore CallSyncCommand(string collectionId, bool conversationMode)
        {
            // Call initial Sync command.
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(collectionId);

            SyncStore syncStore = this.CONAdapter.Sync(syncRequest);

            // Verify Sync command response.
            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            if (conversationMode && Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1")
            {
                syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, null, null, true);
            }
            else
            {
                syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, syncStore.SyncKey, null, null, false);
            }

            syncStore = this.CONAdapter.Sync(syncRequest);

            // Verify Sync command response.
            Site.Assert.AreEqual<byte>(
                1,
                syncStore.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            bool checkSyncStore = syncStore.AddElements != null && syncStore.AddElements.Count != 0;
            Site.Assert.IsTrue(checkSyncStore, "The items should be gotten from the Sync command response.");

            this.LatestSyncKey = syncStore.SyncKey;

            return syncStore;
        }

        /// <summary>
        /// Gets an estimate of the number of items in the specific folder.
        /// </summary>
        /// <param name="syncKey">The latest SyncKey.</param> 
        /// <param name="collectionId">The CollectionId of the folder.</param> 
        /// <returns>The response of GetItemEstimate command.</returns>
        protected GetItemEstimateResponse CallGetItemEstimateCommand(string syncKey, string collectionId)
        {
            // Create GetItemEstimate command request.
            Request.GetItemEstimateCollection collection = new Request.GetItemEstimateCollection
            {
                CollectionId = collectionId,
                SyncKey = syncKey
            };

            if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) != "12.1")
            {
                collection.ConversationMode = true;
                collection.ConversationModeSpecified = true;
            }

            GetItemEstimateRequest getItemEstimateRequest = Common.CreateGetItemEstimateRequest(new Request.GetItemEstimateCollection[] { collection });

            GetItemEstimateResponse getItemEstimateResponse = this.CONAdapter.GetItemEstimate(getItemEstimateRequest);

            return getItemEstimateResponse;
        }

        /// <summary>
        /// Move items to the specific folder.
        /// </summary>
        /// <param name="serverIds">The ServerId of the items to move.</param> 
        /// <param name="sourceFolder">The CollectionId of the source folder.</param>
        /// <param name="destinationFolder">The CollectionId of the destination folder.</param> 
        /// <returns>The response of MoveItems command.</returns>
        protected MoveItemsResponse CallMoveItemsCommand(Collection<string> serverIds, string sourceFolder, string destinationFolder)
        {
            // Move the items from sourceFolder to destinationFolder.
            List<Request.MoveItemsMove> moveItems = new List<Request.MoveItemsMove>();
            foreach (string serverId in serverIds)
            {
                Request.MoveItemsMove move = new Request.MoveItemsMove
                {
                    SrcFldId = sourceFolder,
                    DstFldId = destinationFolder,
                    SrcMsgId = serverId
                };

                moveItems.Add(move);
            }

            MoveItemsRequest moveItemsRequest = Common.CreateMoveItemsRequest(moveItems.ToArray());

            // Call MoveItems command to move the items.
            MoveItemsResponse moveItemsResponse = this.CONAdapter.MoveItems(moveItemsRequest);

            Site.Assert.AreEqual<int>(serverIds.Count, moveItemsResponse.ResponseData.Response.Length, "The count of Response element should be {0}, actual {1}.", serverIds.Count, moveItemsResponse.ResponseData.Response.Length);
            foreach (Response.MoveItemsResponse response in moveItemsResponse.ResponseData.Response)
            {
                Site.Assert.AreEqual<int>(3, int.Parse(response.Status), "If the MoveItems command executes successfully, the Status should be 3, actual {0}.", response.Status);
            }

            return moveItemsResponse;
        }

        /// <summary>
        /// Gets the created ConversationItem.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the parent folder which has the conversation.</param>
        /// <param name="conversationId">The ConversationId of the conversation.</param>
        /// <returns>A ConversationItem object.</returns>
        protected ConversationItem GetConversationItem(string collectionId, string conversationId)
        {
            // Call Sync command to get the emails in Inbox folder.
            SyncStore syncStore = this.CallSyncCommand(collectionId, false);

            // Get the emails from Sync response according to the ConversationId.
            ConversationItem conversationItem = new ConversationItem { ConversationId = conversationId };

            foreach (Sync addElement in syncStore.AddElements)
            {
                if (addElement.Email.ConversationId == conversationId)
                {
                    conversationItem.ServerId.Add(addElement.ServerId);
                }
            }

            Site.Assert.AreNotEqual<int>(0, conversationItem.ServerId.Count, "The conversation should have at least one email.");

            return conversationItem;
        }

        /// <summary>
        /// Gets the conversation with the expected emails count.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the parent folder which has the conversation.</param>
        /// <param name="conversationId">The ConversationId of the conversation.</param>
        /// <param name="expectEmailCount">The expect count of the conversation emails.</param>
        /// <returns>A ConversationItem object.</returns>
        protected ConversationItem GetConversationItem(string collectionId, string conversationId, int expectEmailCount)
        {
            ConversationItem coversationItem;
            int counter = 0;
            int retryLimit = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            do
            {
                System.Threading.Thread.Sleep(waitTime);
                coversationItem = this.GetConversationItem(collectionId, conversationId);
                counter++;
            }
            while (coversationItem.ServerId.Count != expectEmailCount && counter < retryLimit);

            return coversationItem;
        }

        /// <summary>
        /// Checks if ActiveSync Protocol Version is not "14.0".
        /// </summary>
        protected void CheckActiveSyncVersionIsNot140()
        {
            Site.Assume.AreNotEqual<string>("14.0", Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "The airsyncbase:BodyPartPreference element is not supported when the MS-ASProtocolVersion header is set to 14.0. MS-ASProtocolVersion header value is determined using Common PTFConfig property named ActiveSyncProtocolVersion.");
        }
        #endregion

        #region Capture code
        /// <summary>
        /// Verify the message part when the request contains neither BodyPreference nor BodyPartPreference elements.
        /// </summary>
        /// <param name="email">The email item server returned.</param>
        protected void VerifyMessagePartWithoutPreference(Email email)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R245");

            // Verify MS-ASCON requirement: MS-ASCON_R245
            bool isVerifiedR245 = email.Body != null && email.BodyPart == null;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR245,
                245,
                @"[In Sending a Message Part] If request contains neither airsyncbase:BodyPreference nor airsyncbase:BodyPartPreference elements, then the response contains only airsyncbase:Body element.");
        }

        /// <summary>
        /// Verify the message part when the request contains only BodyPreference element.
        /// </summary>
        /// <param name="email">The email item server returned.</param>
        protected void VerifyMessagePartWithBodyPreference(Email email)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R246");

            // Verify MS-ASCON requirement: MS-ASCON_R246
            bool isVerifiedR246 = email.Body != null && email.BodyPart == null;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR246,
                246,
                @"[In Sending a Message Part] If request contains only airsyncbase:BodyPreference element, then the response contains only airsyncbase:Body element.");
        }

        /// <summary>
        /// Verify the message part when the request contains only BodyPartPreference element.
        /// </summary>
        /// <param name="email">The email item server returned.</param>
        /// <param name="truncatedData">The truncated email data returned in BodyPart.</param>
        /// <param name="allData">All email data without being truncated.</param>
        /// <param name="truncationSize">The TruncationSize element specified in BodyPartPreference.</param>
        protected void VerifyMessagePartWithBodyPartPreference(Email email, string truncatedData, string allData, int truncationSize)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R239");

            // Verify MS-ASCON requirement: MS-ASCON_R239
            bool isVerifiedR239 = email.BodyPart.TruncatedSpecified && email.BodyPart.Truncated
                && truncatedData == TestSuiteHelper.TruncateData(allData, truncationSize);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR239,
                239,
                @"[In Sending a Message Part] The client's preferences affect the server response as follows: If the size of the message part exceeds the value specified in the airsyncbase:TruncationSize element ([MS-ASAIRS] section 2.2.2.40.1) of the request, then the server truncates the message part.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R240");

            // Verify MS-ASCON requirement: MS-ASCON_R240
            bool isVerifiedR240 = email.BodyPart.TruncatedSpecified && email.BodyPart.Truncated && email.BodyPart.EstimatedDataSize > 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR240,
                240,
                @"[In Sending a Message Part] The server includes the airsyncbase:Truncated element ([MS-ASAIRS] section 2.2.2.39.1) and the airsyncbase:EstimatedDataSize element ([MS-ASAIRS] section 2.2.2.23.2) in the response when it truncates the message part.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R247");

            // Verify MS-ASCON requirement: MS-ASCON_R247
            bool isVerifiedR247 = email.Body == null && email.BodyPart != null;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR247,
                247,
                @"[In Sending a Message Part] If request contains only airsyncbase:BodyPartPreference element, then the response contains only airsyncbase:BodyPart element.");
        }

        /// <summary>
        /// Verify the message part when the request contains both BodyPreference and BodyPartPreference elements.
        /// </summary>
        /// <param name="email">The email item server returned.</param>
        protected void VerifyMessagePartWithBothPreference(Email email)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R248");

            // Verify MS-ASCON requirement: MS-ASCON_R248
            bool isVerifiedR248 = email.Body != null && email.BodyPart != null;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR248,
                248,
                @"[In Sending a Message Part] If request contains both airsyncbase:BodyPreference and airsyncbase:BodyPartPreference element, then the response contains both airsyncbase:Body and airsyncbase:BodyPart element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R243");

            // Verify MS-ASCON requirement: MS-ASCON_R243
            // If R248 is captured, then BodyPart element and Body element do co-exist in the server response.
            Site.CaptureRequirement(
                243,
                @"[In Sending a Message Part] The airsyncbase:BodyPart element and the airsyncbase:Body element ([MS-ASAIRS] section 2.2.2.9) can co-exist in the server response.");
        }

        /// <summary>
        /// Verify status 164 is returned when the Type element in the BodyPartPreference is other than 2.
        /// </summary>
        /// <param name="status">The status that server returned.</param>
        protected void VerifyMessagePartStatus164(int status)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCON_R241");

            // Verify MS-ASCON requirement: MS-ASCON_R241
            Site.CaptureRequirementIfAreEqual<int>(
                164,
                status,
                241,
                @"[In Sending a Message Part] [The client's preferences affect the server response as follows:] If a value other than 2 is specified in the airsyncbase:Type element ([MS-ASAIRS] section 2.2.2.41.3) of the request, then the server returns a status value of 164.");
        }
        #endregion

        #region Private methods

        /// <summary>
        /// Delete all the items in a folder.
        /// </summary>
        /// <param name="createdItems">The created items which should be deleted.</param>
        private void DeleteItemsInFolder(Collection<CreatedItems> createdItems)
        {
            foreach (CreatedItems createdItem in createdItems)
            {
                SyncStore syncResult = this.CallSyncCommand(createdItem.CollectionId, false);
                List<Request.SyncCollectionDelete> deleteData = new List<Request.SyncCollectionDelete>();
                List<string> serverIds = new List<string>();

                foreach (string subject in createdItem.ItemSubject)
                {
                    if (syncResult != null)
                    {
                        foreach (Sync item in syncResult.AddElements)
                        {
                            if (item.Email.Subject != null && item.Email.Subject.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverIds.Add(item.ServerId);
                            }

                            if (item.Calendar.Subject != null && item.Calendar.Subject.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverIds.Add(item.ServerId);
                            }
                        }
                    }

                    Site.Assert.AreNotEqual<int>(0, serverIds.Count, "The items with subject '{0}' should be found!", subject);

                    foreach (string serverId in serverIds)
                    {
                        deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
                    }

                    Request.SyncCollection syncCollection = new Request.SyncCollection
                    {
                        Commands = deleteData.ToArray(),
                        DeletesAsMoves = false,
                        DeletesAsMovesSpecified = true,
                        CollectionId = createdItem.CollectionId,
                        SyncKey = syncResult.SyncKey
                    };

                    SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                    SyncStore deleteResult = this.CONAdapter.Sync(syncRequest);

                    Site.Assert.AreEqual<byte>(
                    1,
                    deleteResult.CollectionStatus,
                    "The value of Status should be 1 to indicate the Sync command executed successfully.");
                }
            }
        }

        #endregion
    }
}