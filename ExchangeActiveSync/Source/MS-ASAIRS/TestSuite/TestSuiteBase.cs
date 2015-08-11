namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets protocol Interface of MS-ASAIRS
        /// </summary>
        protected IMS_ASAIRSAdapter ASAIRSAdapter { get; private set; }

        /// <summary>
        /// Gets or sets the information of User1.
        /// </summary>
        protected UserInformation User1Information { get; set; }

        /// <summary>
        /// Gets or sets the information of User2.
        /// </summary>
        protected UserInformation User2Information { get; set; }

        /// <summary>
        /// Gets the value of syncKey.
        /// </summary>
        protected string SyncKey { get; private set; }
        #endregion

        #region Test case initialize and cleanup
        /// <summary>
        /// Override the base TestInitialize function
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            if (this.ASAIRSAdapter == null)
            {
                this.ASAIRSAdapter = this.Site.GetAdapter<IMS_ASAIRSAdapter>();
            }

            // Get the information of User1
            this.User1Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Get the information of User2.
            this.User2Information = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                // Switch user to call active sync operations and resynchronize the collection hierarchy.
                this.SwitchUser(this.User1Information, true);
            }
        }

        /// <summary>
        /// Override the base TestCleanup function
        /// </summary>
        protected override void TestCleanup()
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

            base.TestCleanup();
        }
        #endregion

        #region Switch user to another one
        /// <summary>
        /// Change user to call active sync operations and resynchronize the collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of the user that will switch to.</param>
        /// <param name="isFolderSyncNeeded">A boolean value indicates whether needs to synchronize the folder hierarchy.</param>
        protected void SwitchUser(UserInformation userInformation, bool isFolderSyncNeeded)
        {
            this.ASAIRSAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (isFolderSyncNeeded)
            {
                FolderSyncResponse folderSyncResponse = this.FolderSync();

                // Get the folder collectionId of User1
                if (userInformation.UserName == this.User1Information.UserName)
                {
                    if (string.IsNullOrEmpty(this.User1Information.InboxCollectionId))
                    {
                        this.User1Information.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                    }

                    if (string.IsNullOrEmpty(this.User1Information.ContactsCollectionId))
                    {
                        this.User1Information.ContactsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Contacts, this.Site);
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
            }
        }
        #endregion

        #region Get initial SyncKey
        /// <summary>
        /// Get the initial syncKey of the specified folder.
        /// </summary>
        /// <param name="collectionId">The collection id of the specified folder.</param>
        /// <returns>The initial syncKey of the specified folder.</returns>
        protected string GetInitialSyncKey(string collectionId)
        {
            // Obtains the key by sending an initial Sync request with a SyncKey element value of zero and the CollectionId element
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(collectionId);
            DataStructures.SyncStore syncResult = this.ASAIRSAdapter.Sync(syncRequest);

            // Status code '12' means the folder hierarchy has changed
            while (syncResult.Status == 12)
            {
                // Resynchronize the folder hierarchy
                this.FolderSync();

                syncResult = this.ASAIRSAdapter.Sync(syncRequest);
            }

            this.Site.Assert.IsNotNull(
                syncResult,
                "The result for an initial synchronize should not be null.");

            // Verify Sync result
            this.Site.Assert.AreEqual<byte>(
                1,
                syncResult.CollectionStatus,
                "If the Sync command executes successfully, the Status in response should be 1.");

            return syncResult.SyncKey;
        }
        #endregion

        #region Get the specified item
        /// <summary>
        /// Get the non-truncated item data.
        /// </summary>
        /// <param name="subject">The subject of the item.</param>
        /// <param name="collectionId">The server id of the folder which contains the specified item.</param>
        /// <returns>The item with non-truncated data.</returns>
        protected DataStructures.Sync GetAllContentItem(string subject, string collectionId)
        {
            Request.BodyPartPreference[] bodyPartPreference = null;

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
            {
                // Set the BodyPartPreference element to retrieve the BodyPart element in response
                bodyPartPreference = new Request.BodyPartPreference[]
                {
                    new Request.BodyPartPreference()
                    {
                        Type = 2
                    }
                };
            }

            // Set the BodyPreference element to retrieve the Body element in response
            Request.BodyPreference[] bodyPreference = new Request.BodyPreference[]
            {
                new Request.BodyPreference()
                {
                    Type = 1
                }
            };

            // Get the item with specified subject
            DataStructures.Sync item = this.GetSyncResult(subject, collectionId, null, bodyPreference, bodyPartPreference);

            this.Site.Assert.IsNotNull(item.Email.Body, "The Body element should be included in Sync command response when the BodyPreference element is specified in Sync command request.");

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
            {
                this.Site.Assert.IsNotNull(
                    item.Email.BodyPart,
                    "The BodyPart element should be included in Sync command response when the BodyPartPreference element is specified in Sync command request.");

                this.Site.Assert.AreEqual<byte>(
                    1,
                    item.Email.BodyPart.Status,
                    "The Status should be 1 to indicate the success of the Sync command response in returning Data element content given the BodyPartPreference element settings in the Sync command request.");
            }

            return item;
        }

        /// <summary>
        /// Synchronize item with specified subject.
        /// </summary>
        /// <param name="subject">The subject of the item.</param>
        /// <param name="collectionId">The collection id which to sync with.</param>
        /// <param name="commands">The sync commands.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>The item with specified subject.</returns>
        protected DataStructures.Sync GetSyncResult(string subject, string collectionId, object[] commands, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            DataStructures.SyncStore syncStore;
            DataStructures.Sync item = null;
            SyncRequest request = TestSuiteHelper.CreateSyncRequest(this.GetInitialSyncKey(collectionId), collectionId, commands, bodyPreferences, bodyPartPreferences);

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);
                syncStore = this.ASAIRSAdapter.Sync(request);
                if (syncStore != null && syncStore.CollectionStatus == 1)
                {
                    item = TestSuiteHelper.GetSyncAddItem(syncStore, subject);
                }

                counter++;
            }
            while ((syncStore == null || item == null) && counter < retryCount);

            this.Site.Assert.IsNotNull(item, "The email item with subject {0} should be found, retry count: {1}.", subject, counter);

            this.SyncKey = syncStore.SyncKey;

            return item;
        }

        /// <summary>
        /// Fetch item with specified ServerID on the server.
        /// </summary>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="serverId">The server id of the mail.</param>
        /// <param name="fileReference">The file reference of the attachment.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <param name="deliveryMethod">Indicate whether use multipart or inline method to send the request.</param>
        /// <returns>The item with specified ServerID.</returns>
        protected DataStructures.ItemOperations GetItemOperationsResult(string collectionId, string serverId, string fileReference, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences, DeliveryMethodForFetch? deliveryMethod)
        {
            DataStructures.ItemOperations item = null;
            ItemOperationsRequest request = TestSuiteHelper.CreateItemOperationsRequest(collectionId, serverId, fileReference, bodyPreferences, bodyPartPreferences);

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                Thread.Sleep(waitTime);
                DataStructures.ItemOperationsStore itemOperationsStore = this.ASAIRSAdapter.ItemOperations(request, deliveryMethod ?? DeliveryMethodForFetch.Inline);

                // Since the item serverId or attachment fileReference is unique, there should be only one item in response
                this.Site.Assert.AreEqual<int>(
                    1,
                    itemOperationsStore.Items.Count,
                    "The count of Items in ItemOperations command response should be 1.");

                if (itemOperationsStore.Items[0].Email != null)
                {
                    item = itemOperationsStore.Items[0];
                }
            }
            while (item == null && counter < retryCount);

            this.Site.Assert.IsNotNull(item, "The item should be found, retry count: {0}.", counter);

            return item;
        }

        /// <summary>
        /// Search item with specified criteria on the server.
        /// </summary>
        /// <param name="subject">The subject of the item.</param>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="conversationId">The conversation for which to search.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>The server response.</returns>
        protected DataStructures.Search GetSearchResult(string subject, string collectionId, string conversationId, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            SearchRequest request = TestSuiteHelper.CreateSearchRequest(subject, collectionId, conversationId, bodyPreferences, bodyPartPreferences);

            DataStructures.SearchStore searchStore = this.ASAIRSAdapter.Search(request);
            DataStructures.Search searchItem = null;
            if (searchStore.Results.Count != 0)
            {
                searchItem = TestSuiteHelper.GetSearchItem(searchStore, subject);
            }

            this.Site.Assert.IsNotNull(searchItem, "The email message with subject {0} should be found.", subject);

            return searchItem;
        }
        #endregion

        #region Send mail from user1 to user2
        /// <summary>
        /// The method is used to send a mail
        /// </summary>
        /// <param name="emailType">The type of the email.</param>
        /// <param name="subject">The subject of the mail.</param>
        /// <param name="body">The body of the item.</param>
        protected void SendEmail(EmailType emailType, string subject, string body)
        {
            SendMailRequest request = new SendMailRequest
            {
                RequestData =
                {
                    ClientId = Guid.NewGuid().ToString("N"),
                    Mime = TestSuiteHelper.CreateMIME(
                        emailType,
                        Common.GetMailAddress(this.User1Information.UserName, this.User1Information.UserDomain),
                        Common.GetMailAddress(this.User2Information.UserName, this.User2Information.UserDomain),
                        subject,
                        body)
                }
            };

            SendMailResponse sendMailResponse = this.ASAIRSAdapter.SendMail(request);
            this.Site.Assert.AreEqual<string>(
                 string.Empty,
                 sendMailResponse.ResponseDataXML,
                 "The server should return an empty XML body to indicate SendMail command is executed successfully.");

            this.SwitchUser(this.User2Information, true);
            this.RecordCaseRelativeItems(this.User2Information.UserName, this.User2Information.InboxCollectionId, subject);
        }
        #endregion

        #region Record the userName, folder collectionId and item subject
        /// <summary>
        /// Record the user name, folder collectionId and subjects the current test case impacts.
        /// </summary>
        /// <param name="userName">The user that current test case used.</param>
        /// <param name="folderCollectionId">The collectionId of folders that the current test case impact.</param>
        /// <param name="itemSubjects">The subject of items that the current test case impact.</param>
        protected void RecordCaseRelativeItems(string userName, string folderCollectionId, params string[] itemSubjects)
        {
            // Record the item in the specified folder.
            CreatedItems createdItems = new CreatedItems { CollectionId = folderCollectionId };

            foreach (string subject in itemSubjects)
            {
                createdItems.ItemSubject.Add(subject);
            }

            // Record the created items of User1.
            if (userName == this.User1Information.UserName)
            {
                this.User1Information.UserCreatedItems.Add(createdItems);
            }

            // Record the created items of User2.
            if (userName == this.User2Information.UserName)
            {
                this.User2Information.UserCreatedItems.Add(createdItems);
            }
        }
        #endregion

        #region Private methods
        /// <summary>
        /// This method is used to synchronize the folder collection hierarchy.
        /// </summary>
        /// <returns>The response of the FolderSync command.</returns>
        private FolderSyncResponse FolderSync()
        {
            FolderSyncResponse folderSyncResponse = this.ASAIRSAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

            this.Site.Assert.AreEqual<byte>(
                1,
                folderSyncResponse.ResponseData.Status,
                "The Status value should be 1 to indicate the FolderSync command executes successfully.");

            return folderSyncResponse;
        }

        /// <summary>
        /// Delete all the items in a folder.
        /// </summary>
        /// <param name="createdItemsCollection">The created items collection which should be deleted.</param>
        private void DeleteItemsInFolder(Collection<CreatedItems> createdItemsCollection)
        {
            foreach (CreatedItems createdItems in createdItemsCollection)
            {
                string syncKey = this.GetInitialSyncKey(createdItems.CollectionId);
                SyncRequest request = TestSuiteHelper.CreateSyncRequest(syncKey, createdItems.CollectionId, null, null, null);
                DataStructures.SyncStore result = this.ASAIRSAdapter.Sync(request);

                List<Request.SyncCollectionDelete> deleteData = new List<Request.SyncCollectionDelete>();
                foreach (string subject in createdItems.ItemSubject)
                {
                    string serverId = null;
                    if (result != null)
                    {
                        foreach (DataStructures.Sync item in result.AddElements)
                        {
                            if (item.Email.Subject != null && item.Email.Subject.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverId = item.ServerId;
                                break;
                            }

                            if (item.Contact.FileAs != null && item.Contact.FileAs.Equals(subject, StringComparison.CurrentCulture))
                            {
                                serverId = item.ServerId;
                                break;
                            }
                        }
                    }

                    this.Site.Assert.IsNotNull(serverId, "The item with subject '{0}' should be found!", subject);
                    deleteData.Add(new Request.SyncCollectionDelete() { ServerId = serverId });
                }

                Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(result.SyncKey, createdItems.CollectionId);
                syncCollection.Commands = deleteData.ToArray();
                syncCollection.DeletesAsMoves = false;
                syncCollection.DeletesAsMovesSpecified = true;

                SyncRequest syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                DataStructures.SyncStore deleteResult = this.ASAIRSAdapter.Sync(syncRequest);
                this.Site.Assert.AreEqual<byte>(
                    1,
                    deleteResult.CollectionStatus,
                    "The value of Status should be 1 to indicate the Sync command executes successfully.");
            }
        }
        #endregion
    }
}