namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test SyncFolderItems operation on the following items: MessageType item, MeetingRequestMessageType item, MeetingResponseMessageType item, MeetingCancellationMessageType item, TaskType item, ContactItemType item, PostItemType item, CalendarItemType item, DistributionListType item and ItemType item.
    /// </summary>
    [TestClass]
    public class S02_SyncFolderItems : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test cases
        /// <summary>
        /// Client calls SyncFolderItems operation to sync MessageType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC01_SyncFolderItems_MessageType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get the initial syncState of inbox folder.
            DistinguishedFolderIdNameType inboxFolder = DistinguishedFolderIdNameType.inbox;

            // Set DefaultShapeNamesType to AllProperties
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(inboxFolder, DefaultShapeNamesType.AllProperties);

            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem operation to create a MessageType item on server and get its id.
            MessageType messageType = new MessageType();
            BaseItemIdType[] itemIds = this.CreateItem(inboxFolder, messageType);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R70");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R70
            Site.CaptureRequirementIfIsNotNull(
                changes,
                70,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] [The element Changes] specifies the differences between the items on the client and the items on the server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MessageType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is MessageType, then requirement MS-OXWSSYNC_R156 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R156");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R156
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(MessageType),
                156,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of Message is t:MessageType ([MS-OXWSMSG] section 2.2.4.1).");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MessageType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isMessageTypeItemCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MessageType);

            // If the ItemsElementName of Changes is Create and the type of Item is MessageType, it indicates a message has been created on server and synced on client, 
            // then requirements MS-OXWSSYNC_R131 and MS-OXWSSYNC_R1571 can be captured
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R131. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(MessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R131
            Site.CaptureRequirementIfIsTrue(
                isMessageTypeItemCreated,
                131,
                @"[In t:SyncFolderItemsChangesType Complex Type] [The element Create] specifies an item that has been created on the server and has to be created on the client.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1571. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(MessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1571
            Site.CaptureRequirementIfIsTrue(
                isMessageTypeItemCreated,
                1571,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element Message] specifies a message to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the subject of the item that created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, inboxFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState  to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isItemSubjectUpdated = false;
            isItemSubjectUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update &&
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.Subject == newItemSubject;

            // If the ItemsElementName of Changes is Update and the item's subject is a new value, it indicates the message type item has been updated on server and synced on client.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R132. Expected value: ItemsElementName: {0}, item subject: {1}; actual value: ItemsElementName: {2}, item subject: {3}",
                ItemsChoiceType1.Update,
                newItemSubject,
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.Subject);

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R132
            Site.CaptureRequirementIfIsTrue(
                isItemSubjectUpdated,
                132,
                @"[In t:SyncFolderItemsChangesType Complex Type] [The element Update] specifies an item that has been changed on the server and has to be changed on the client.");

            bool isMessageUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update &&
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MessageType);

            // If the ItemsElementName of Changes is Update and the type of Item is MessageType, it indicates a message has been updated on server and synced on client.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1572. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(MessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1572
            Site.CaptureRequirementIfIsTrue(
                isMessageUpdated,
                1572,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element Message] specifies a message to  update in the client data store.");
            #endregion

            #region Step 6. Client invokes UpdateItem operation to change the IsRead property of the item that updated in Step 4.
            // Call UpdateReadFlag to update the IsRead property.
            this.UpdateReadFlag(itemIds, this.ConvertReadFlag(itemIds));
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isUpdateReadFlag = changes.ItemsElementName[0] == ItemsChoiceType1.ReadFlagChange && changes.Items[0].GetType() == typeof(SyncFolderItemsReadFlagType);

            // If the ItemsElementName of changes is ReadFlagChange and the type of item in changes is SyncFolderItemsReadFlagType, 
            // it indicates the read flag has been updated on server and synced on client.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R136. Expected value: ItemsElementName: {0}, change items type: {1}; actual value: ItemsElementName: {2}, change items type: {3}",
                ItemsChoiceType1.ReadFlagChange,
                typeof(SyncFolderItemsReadFlagType),
                changes.ItemsElementName[0],
                changes.Items[0].GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R136
            Site.CaptureRequirementIfIsTrue(
                isUpdateReadFlag,
                136,
                @"[In t:SyncFolderItemsChangesType Complex Type] [The element ReadFlagChange] specifies an item that has been marked as read on the server and has to be marked as read on the client.");
            #endregion

            #region Step 8. Client invokes DeleteItem operation to delete the item which the IsRead property is updated in Step 6.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 9. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 8 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MessageType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MessageType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isItemDeleted = (changes.ItemsElementName[0] == ItemsChoiceType1.Delete) && (changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType));

            // If the ItemsElementName is Delete and the items type in changes is SyncFolderItemsDeleteType, it indicates a MessageType item has been deleted on server and synced on client. 
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R134. Expected value: ItemsElementName: {0}, change items type: {1}; actual value: ItemsElementName: {2}, change items type: {3}",
                ItemsChoiceType1.Delete,
                typeof(SyncFolderItemsDeleteType),
                changes.ItemsElementName[0],
                changes.Items[0].GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R134
            Site.CaptureRequirementIfIsTrue(
                isItemDeleted,
                134,
                @"[In t:SyncFolderItemsChangesType Complex Type] [The element Delete] specifies an item that has been deleted on the server and has to be deleted on the client.");
            #endregion
        }

        /// <summary>
        ///  Client calls SyncFolderItems operation to sync MeetingRequestMessageType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC02_SyncFolderItems_MeetingRequestMessageType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get the initial syncState of sent items folder.
            DistinguishedFolderIdNameType sentItemsFolder = DistinguishedFolderIdNameType.sentitems;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(sentItemsFolder, DefaultShapeNamesType.Default);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem operation to create a MeetingRequestMessageType item.
            // Generate the item subject
            string itemSubject = Common.GenerateResourceName(this.Site, sentItemsFolder + "ItemSubject");
            this.CreateMeetingRequest(this.User2EmailAddress, itemSubject);

            // Make sure that the meeting request exists in inbox folder of User2.
            bool isReceivedInInbox = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.inbox.ToString(),
                itemSubject,
                Item.MeetingRequest.ToString());
            Site.Assert.IsTrue(
                isReceivedInInbox,
                string.Format("The meeting request message should exist in inbox folder of '{0}'", Common.GetConfigurationPropertyValue("User2Name", this.Site)));
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(sentItemsFolder, responseMessage, DefaultShapeNamesType.Default);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingRequestMessageType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is MeetingRequestMessageType, then requirement MS-OXWSSYNC_R166 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R166");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R166
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(MeetingRequestMessageType),
                166,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of MeetingRequest is t:MeetingRequestMessageType ([MS-OXWSMTGS] section 2.2.4.13).");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingRequestMessageType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isMeetingRequestCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingRequestMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1671. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(MeetingRequestMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is MeetingRequestMessageType, it indicates a meeting 
            // request has been created on server and synced on client, then requirement MS-OXWSSYNC_R1671 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1671
            Site.CaptureRequirementIfIsTrue(
                isMeetingRequestCreated,
                1671,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element MeetingRequest] specifies a meeting request message to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the subject of the item that created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, sentItemsFolder + "NewItemSubject");
            ItemIdType[] itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            this.UpdateItemSubject(itemId, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to get the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(sentItemsFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingRequestMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingRequestMessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isMeetingRequestUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update &&
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingRequestMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1672. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(MeetingRequestMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is MeetingRequestMessageType, it indicates a meeting 
            // request has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1672 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1672
            Site.CaptureRequirementIfIsTrue(
                isMeetingRequestUpdated,
                1672,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element MeetingRequest] specifies a meeting request message to update in the client data store.");

            // Call GetItem operation to get the parent folder Id of the item that in SyncFolderItems response.
            GetItemType getItemRequest = new GetItemType();
            getItemRequest.ItemIds = new BaseItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            getItemRequest.ItemShape = new ItemResponseShapeType();
            getItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            GetItemResponseType getItemResponse = this.COREAdapter.GetItem(getItemRequest);

            // Check whether the GetItem operation is executed successfully.
            foreach (ResponseMessageType message in getItemResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        message.ResponseClass,
                        string.Format("Get item should be successful! Expected response code: {0}, actual response code: {1}", ResponseCodeType.NoError, message.ResponseCode));
            }

            // Get the item information from GetItem response
            MeetingRequestMessageType[] item = Common.GetItemsFromInfoResponse<MeetingRequestMessageType>(getItemResponse);
            Site.Assert.AreEqual<int>(1, item.Length, "Only one MeetingRequestMessageType item was created, so there should be only 1 item in GetItem response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R380");

            // If the parent folder Id of the item in GetItem response is same with it in SyncFolderItems response, 
            // it indicates the folder contains the item to synchronize, then requirement MS-OXWSSYNC_R380 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R380
            Site.CaptureRequirementIfAreEqual<string>(
                item[0].ParentFolderId.Id,
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ParentFolderId.Id,
                380,
                @"[In m:SyncFolderItemsType Complex Type] [The element SyncFolderId] specifies the identity of the folder that contains the items to synchronize.");
            #endregion

            #region Step 6. Client invokes UpdateItem operation to change the IsRead property of the item which updated in Step 4.
            itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            this.UpdateReadFlag(itemId, this.ConvertReadFlag(itemId));
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(sentItemsFolder, responseMessage, DefaultShapeNamesType.Default);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changesAfterUpdateReadFlag = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changesAfterUpdateReadFlag.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changesAfterUpdateReadFlag.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changesAfterUpdateReadFlag.Items.Length, "Just one MeetingRequestMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            #endregion

            #region Step 8. Client invokes DeleteItem operation to delete the item which updated in Step 6.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteItem(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                sentItemsFolder.ToString(),
                newItemSubject,
                Item.MeetingRequest.ToString());
            Site.Assert.IsTrue(isDeleted, string.Format("The item named '{0}' should be deleted from folder '{1}' successfully.", newItemSubject, sentItemsFolder));
            #endregion

            #region Step 9. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 8 and verify related requirements.
            responseMessage = this.GetResponseMessage(sentItemsFolder, responseMessage, DefaultShapeNamesType.Default);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changesAfterDelete = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changesAfterDelete.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changesAfterDelete.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changesAfterDelete.Items.Length, "Just one MeetingRequestMessageType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changesAfterDelete.ItemsElementName.Length, "Just one MeetingRequestMessageType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R184");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R184
            Site.CaptureRequirementIfAreEqual<string>(
                (changesAfterUpdateReadFlag.Items[0] as SyncFolderItemsReadFlagType).ItemId.Id,
                (changesAfterDelete.Items[0] as SyncFolderItemsDeleteType).ItemId.Id,
                184,
                @"[In t:SyncFolderItemsDeleteType Complex Type] [The element ItemId] specifies the identifier of the item to delete from the client data store.");

            bool isIncrementalSync = changesAfterDelete.ItemsElementName[0] == ItemsChoiceType1.Delete && responseMessage.SyncState != null;

            // If the ItemsElementName is Delete and the SyncState element is not null, then requirement MS-OXWSSYNC_R504 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R504");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R504
            Site.CaptureRequirementIfIsTrue(
                isIncrementalSync,
                504,
                @"[In Abstract Data Model]  If the SyncState element of the SyncFolderItemsType complex type (section 3.1.4.2.3.1) is included in a SyncFolderItems operation (section 3.1.4.2), the server MUST return incremental synchronization information from the last synchronization request. ");
            #endregion

            #region Step 10 Clean up the mailbox of attendee.
            this.CleanupAttendeeMailbox();
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync MeetingResponseMessageType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC03_SyncFolderItems_MeetingResponseMessageType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get the initial syncState of inbox folder.
            DistinguishedFolderIdNameType inboxFolder = DistinguishedFolderIdNameType.inbox;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(inboxFolder, DefaultShapeNamesType.IdOnly);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3832");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3832
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                3832,
                @"[In m:SyncFolderItemsType Complex Type] This element [SyncState] not present, server responses NO_ERROR.");
            #endregion

            #region Step 2. Client invokes CreateMeetingResponse operation to create a MeetingResponseMessageType item.
            // Generate the item subject
            string itemSubject = Common.GenerateResourceName(this.Site, inboxFolder + "ItemSubject");
            this.CreateMeetingResponse(itemSubject);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to get the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.IdOnly);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3831");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3831
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                3831,
                @"[In m:SyncFolderItemsType Complex Type] This element [SyncState] is present, server responses NO_ERROR.");

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingResponseMessageType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is MeetingResponseMessageType, then requirement MS-OXWSSYNC_R168 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R168");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R168
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(MeetingResponseMessageType),
                168,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of MeetingResponse is t:MeetingResponseMessageType ([MS-OXWSMTGS] section 2.2.4.14).");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingResponseMessageType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isMeetingResponseCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingResponseMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1691. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(MeetingResponseMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is MeetingResponseMessageType, it indicates a meeting 
            // response has been created on server and synced on client, then requirement MS-OXWSSYNC_R1691 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1691
            Site.CaptureRequirementIfIsTrue(
                isMeetingResponseCreated,
                1691,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element MeetingResponse] specifies a meeting response message to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, inboxFolder + "NewItemSubject");
            ItemIdType[] itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            this.UpdateItemSubject(itemId, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.IdOnly);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingResponseMessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingResponseMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            bool isMeetingResponseUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update &&
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingResponseMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1692. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(MeetingResponseMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is MeetingResponseMessageType, it indicates a meeting 
            // response has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1692 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1692
            Site.CaptureRequirementIfIsTrue(
                isMeetingResponseUpdated,
                1692,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element MeetingResponse] specifies a meeting response message to update in the client data store.");

            bool isIdOnly = Common.IsIdOnly((XmlElement)this.SYNCAdapter.LastRawResponseXml, "t:MeetingResponse", "t:ItemId");

            // If there is only a ItemId element in the item of changes, then requirement MS-OXWSSYNC_R3783 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3783");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3783
            Site.CaptureRequirementIfIsTrue(
                isIdOnly,
                3783,
                @"[In m:SyncFolderItemsType Complex Type] ItemShape element BaseShape, value=IdOnly, specifies only the item or folder ID.");
            #endregion

            #region Step 6. Client invokes UpdateItem operation to change the IsRead property of the item which updated in Step 4.
            itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            this.UpdateReadFlag(itemId, this.ConvertReadFlag(itemId));
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.IdOnly);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingResponseMessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingResponseMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            #endregion

            #region Step 8. Invokes DeleteItem operation to delete the item which updated in Step 6.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteItem(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                inboxFolder.ToString(),
                newItemSubject,
                Item.MeetingResponse.ToString());
            Site.Assert.IsTrue(isDeleted, string.Format("The item named '{0}' should be deleted from folder '{1}' successfully.", newItemSubject, inboxFolder));
            #endregion

            #region Step 9. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 8 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.IdOnly);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingResponseMessageType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingResponseMessageType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The SyncState in response should not be null.");

            // If the SyncState element in SyncFolderItems response is not null, it indicates the synchronization state is returned in response.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R65");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R65
            Site.CaptureRequirement(
                65,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] [The element SyncState] specifies a form of the synchronization data, which is encoded with base64 encoding, that is used to identify the synchronization state.");
            #endregion

            #region Step 10 Clean up the mailbox of attendee.
            this.CleanupAttendeeMailbox();
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation sync MeetingCancellationMessageType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC04_SyncFolderItems_MeetingCancellationMessageType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get the initial syncState of junkemail folder.
            DistinguishedFolderIdNameType junkeFolder = DistinguishedFolderIdNameType.junkemail;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(junkeFolder, DefaultShapeNamesType.AllProperties);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem to create a MeetingCancellationMessageType item.
            // Generate the item subject
            string itemSubject = Common.GenerateResourceName(this.Site, junkeFolder + "ItemSubject");
            this.CreateMeetingCancellation(itemSubject);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(junkeFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be folders information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingCancellationMessageType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is MeetingCancellationMessageType, then requirement MS-OXWSSYNC_R170 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R170");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R170
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(MeetingCancellationMessageType),
                170,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of MeetingCancellation is t:MeetingCancellationMessageType ([MS-OXWSMTGS] section 2.2.4.11).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1951");

            // For the item creator, the read flag of a new created item should be true
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1951
            Site.CaptureRequirementIfIsTrue(
                ((changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item as MeetingCancellationMessageType).IsRead,
                1951,
                @"[In t:SyncFolderItemsReadFlagType Complex Type] [The element IsRead] True if the item has been read.");

            // Assert both the length of responseMessage.Changes.ItemsElementName and responseMessage.Changes.Items are 1.
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingCancellationMessageType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isMeetingCancellationCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingCancellationMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1711. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(MeetingCancellationMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is MeetingCancellationMessageType, it indicates a meeting 
            // cancellation has been created on server and synced on client, then requirement MS-OXWSSYNC_R1711 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1711
            Site.CaptureRequirementIfIsTrue(
                isMeetingCancellationCreated,
                1711,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element MeetingCancellation] specifies a meeting cancellation message to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, junkeFolder + "NewItemSubject");
            ItemIdType[] itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
            this.UpdateItemSubject(itemId, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(junkeFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingCancellationMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingCancellationMessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isMeetingCancellationUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update &&
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MeetingCancellationMessageType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1712. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(MeetingCancellationMessageType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is MeetingCancellationMessageType, it indicates a meeting 
            // cancellation has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1712 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1712
            Site.CaptureRequirementIfIsTrue(
                isMeetingCancellationUpdated,
                1712,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element MeetingCancellation] specifies a meeting cancellation message to update in the client data store.");
            #endregion

            #region Step 6. Client invokes UpdateItem operations to change the IsRead property of the item which updated in Step 4.
            itemId = new ItemIdType[1] { (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };

            // Call UpdateReadFlag to update the IsRead property to its opposite value.
            this.UpdateReadFlag(itemId, this.ConvertReadFlag(itemId));
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6 and verify related requirements.
            responseMessage = this.GetResponseMessage(junkeFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changesAfterUpdateReadFlag = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changesAfterUpdateReadFlag.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changesAfterUpdateReadFlag.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changesAfterUpdateReadFlag.Items.Length, "Just one MeetingCancellationMessageType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R194");

            // If the Item id is same before and after the update of read flag, then requirement MS-OXWSSYNC_R194 can be captured 
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R194
            Site.CaptureRequirementIfAreEqual<string>(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId.Id,
                (changesAfterUpdateReadFlag.Items[0] as SyncFolderItemsReadFlagType).ItemId.Id,
                194,
                @"[In t:SyncFolderItemsReadFlagType Complex Type] [The element ItemID] specifies the identifier of the read item.");

            Site.Assert.AreEqual<int>(1, changesAfterUpdateReadFlag.ItemsElementName.Length, "Just one MeetingCancellationMessageType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1952");

            // Since the original value of IsRead is true, if it was updated to an opposite value, it should be false.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1952
            Site.CaptureRequirementIfIsFalse(
                (changesAfterUpdateReadFlag.Items[0] as SyncFolderItemsReadFlagType).IsRead,
                1952,
                @"[In t:SyncFolderItemsReadFlagType Complex Type] [The element IsRead] False if the item hasn't been read.");
            #endregion

            #region Step 8. Client invokes DeleteItem operation to delete the item which updated in Step 6.
            bool isDeleted = this.SYNCSUTControlAdapter.FindAndDeleteItem(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                junkeFolder.ToString(),
                newItemSubject,
                Item.MeetingCancellation.ToString());
            Site.Assert.IsTrue(isDeleted, string.Format("The item named '{0}' should be deleted from '{1}' successfully.", newItemSubject, junkeFolder));
            #endregion

            #region Step 9. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 8 and verify related requirements.
            responseMessage = this.GetResponseMessage(junkeFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one MeetingCancellationMessageType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // Assert the Items is an instance of SyncFolderItemsDeleteType.
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one MeetingCancellationMessageType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(changes.ItemsElementName[0] == ItemsChoiceType1.Delete, string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));
            #endregion

            #region Step 10 Clean up the mailbox of attendee.
            this.CleanupAttendeeMailbox();
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation sync TaskType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC05_SyncFolderItems_TaskType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of tasks folder.
            DistinguishedFolderIdNameType taskFolder = DistinguishedFolderIdNameType.tasks;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(taskFolder, DefaultShapeNamesType.AllProperties);

            // Involve InlineImageUrlTemplate element in SyncFolderItems request
            if (Common.IsRequirementEnabled(37809, this.Site))
            {
                request.ItemShape.InlineImageUrlTemplate = "Test Template";
            }

            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem to create a TaskType item and get its ID.
            TaskType taskItem = new TaskType();

            // Create two items on server to verify if the MaxSyncChangesReturned is set to 1, the last item should not be included in response
            BaseItemIdType[] firstItemId = this.CreateItem(taskFolder, taskItem);
            BaseItemIdType[] secondItemId = this.CreateItem(taskFolder, taskItem);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;

            // Set MaxSyncChangesReturned to a value that is less than the number of changes to be returned to verify the false value of "IncludesLastItemInRange" element
            request.MaxChangesReturned = 1;
            response = this.SYNCAdapter.SyncFolderItems(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R6702");

            // There are two items created in step2, but the MaxSyncChangesReturned is set to 1, so the last item should not be included in response
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R6702
            Site.CaptureRequirementIfIsFalse(
                responseMessage.IncludesLastItemInRange,
                6702,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] [The element IncludesLastItemInRange] False indicates the last item to synchronize is not included in the response.");

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(
                1,
                changes.Items.Length,
                "There are two TaskType items were created in previous step, but if the MaxChangesReturned is set to 1, the count of Items array in responseMessage.Changes should be 1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3891");

            // There are two items created on server, but if the value of MaxChangesRetured is set to 1, there should be just one item returned in changes
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3891
            Site.CaptureRequirement(
                3891,
                @"[In m:SyncFolderItemsType Complex Type] This element [MaxChangesReturned] is set between 1 and 512, inclusive, the correct changes number returned in a synchronization response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R388");

            // If the description of MS-OXWSSYNC_R3891 is true, then requirement MS-OXWSSYNC_R388 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R388
            Site.CaptureRequirement(
                388,
                @"[In m:SyncFolderItemsType Complex Type] [The element MaxChangesReturned] specifies the maximum number of changes that can be returned in a synchronization response.");

            // If the type of item in SyncFolderItems response is TaskType, then requirement MS-OXWSSYNC_R172 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R172");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R172
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(TaskType),
                172,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of Task is t:TaskType ([MS-OXWSTASK] section 2.2.4.3).");

            Site.Assert.AreEqual<int>(
                1,
                changes.ItemsElementName.Length,
                "There are two TaskType items were created in previous step, but if the MaxChangesReturned is set to 1, the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isTaskTypeItemCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(TaskType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1731. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(TaskType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is TaskType, it indicates a task 
            // has been created on server and synced on client, then requirement MS-OXWSSYNC_R1731 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1731
            Site.CaptureRequirementIfIsTrue(
                isTaskTypeItemCreated,
                1731,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element Task] specifies a task to create in the client data store.");

            // If MS-OXWSSYNC_R37809 is enabled, then verify this requirement
            if (Common.IsRequirementEnabled(37809, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R37809");

                // If the request involves the InlineImageUrlTemplate element and response is successful, then requirement MS-OXWSSYNC_37809 can be captured.
                // Verify MS-OXWSSYNC: MS-OXWSSYNC_R37809
                Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    responseMessage.ResponseClass,
                    37809,
                    @"[In Appendix C: Product Behavior] Implementation does use the InlineImageUrlTemplate element. The InlineImageUrlTemplate element which is subelement 
                of ItemShape specifies the name of the template for the inline image URL. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Step 4. Client invokes DeleteItem operation to delete the second item that created in step2.
            this.DeleteItem(secondItemId);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to get the SyncState
            responseMessage = this.GetResponseMessage(taskFolder, responseMessage, DefaultShapeNamesType.AllProperties);
            #endregion

            #region Step 6. Client invokes UpdateItem operation to update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, taskFolder + "NewItemSubject");
            this.UpdateItemSubject(firstItemId, newItemSubject);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6 and verify related requirements.
            responseMessage = this.GetResponseMessage(taskFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one TaskType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one TaskType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isTaskItemUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(TaskType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1732. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(TaskType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is TaskType, it indicates a task
            // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1732 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1732
            Site.CaptureRequirementIfIsTrue(
                isTaskItemUpdated,
                1732,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element Task] specifies a task to update in the client data store.");
            #endregion

            #region Step 8. Client invokes DeleteItem operation to delete the TaskType item which updated in Step 6.
            this.DeleteItem(firstItemId);
            #endregion

            #region Step 9. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 8.
            responseMessage = this.GetResponseMessage(taskFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one TaskType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one TaskType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync ContactItemType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC06_SyncFolderItems_ContactItemType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of contacts folder.
            DistinguishedFolderIdNameType contactFolder = DistinguishedFolderIdNameType.contacts;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(contactFolder, DefaultShapeNamesType.AllProperties);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem to create a ContactItemType item and get its ID.
            ContactItemType contactItem = new ContactItemType();
            ItemIdType[] itemIds = this.CreateItem(contactFolder, contactItem);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            // Call SyncFolderItems operation to verify the Ignore element after creating an item on server
            SyncFolderItemsType requestWithIgnoreElement = this.CreateSyncFolderItemsRequestWithoutOptionalElements(contactFolder, DefaultShapeNamesType.AllProperties);

            // Involve Ignore element in SyncFolderItems request
            requestWithIgnoreElement.SyncState = responseMessage.SyncState;
            requestWithIgnoreElement.Ignore = new ItemIdType[] { itemIds[0] };
            SyncFolderItemsResponseType responseWithIgnoreElement = this.SYNCAdapter.SyncFolderItems(requestWithIgnoreElement);
            SyncFolderItemsResponseMessageType responseMessageWithIgnoreElement = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(responseWithIgnoreElement);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3861");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3861
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessageWithIgnoreElement.ResponseCode,
                3861,
                @"[In m:SyncFolderItemsType Complex Type] This element [Ignore] is present, server responses NO_ERROR.");

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changesWithIgnoreElement = responseMessageWithIgnoreElement.Changes;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify requirement MS-OXWSSYNC_R3851");

            // If the items in Changes in SyncFolderItems response is null, in indicates the synchronization for the item is skipped.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3851
            Site.CaptureRequirementIfIsNull(
                changesWithIgnoreElement.Items,
                3851,
                @"[In m:SyncFolderItemsType Complex Type] If the item is in Ignore array, the synchronization is skipped.");

            // Call SyncFolderItems operation again without Ignore to get the CreateItem operation result
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3862");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3862
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                responseMessage.ResponseCode,
                3862,
                @"[In m:SyncFolderItemsType Complex Type] This element [Ignore] is not present, server responses NO_ERROR.");

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ContactItemType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R3852. Expected item type: {0}, actual item type: {1}",
                typeof(ContactItemType),
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If there is one contact item in Changes in SyncFolderItems response, it indicates the synchronization for the item is not skipped.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3852
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(ContactItemType),
                3852,
                @"[In m:SyncFolderItemsType Complex Type] If the item is not in Ignore array, the synchronization is not skipped.");

            // If the type of item in SyncFolderItems response is ContactItemType, then requirement MS-OXWSSYNC_R160 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R160");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R160
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(ContactItemType),
                160,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of Contact is t:ContactItemType ([MS-OXWSCONT] section 2.2.4.2).");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ContactItemType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isContactItemCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(ContactItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1611. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(ContactItemType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is ContactItemType, it indicates a contact 
            // has been created on server and synced on client, then requirement MS-OXWSSYNC_R1611 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1611
            Site.CaptureRequirementIfIsTrue(
                isContactItemCreated,
                1611,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element Contact] specifies a contact to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, contactFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ContactItemType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ContactItemType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isContactItemUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(ContactItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1612. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(ContactItemType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is TaskType, it indicates a task
            // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1612 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1612
            Site.CaptureRequirementIfIsTrue(
                isContactItemUpdated,
                1612,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element Contact] specifies a contact to update in the client data store.");

            // Call SyncFolderItems again without SyncState to verify that all synchronization is returned.
            SyncFolderItemsType requestWithoutSyncState = this.CreateSyncFolderItemsRequestWithoutOptionalElements(contactFolder, DefaultShapeNamesType.AllProperties);
            response = this.SYNCAdapter.SyncFolderItems(requestWithoutSyncState);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<ItemsChoiceType1>(
                ItemsChoiceType1.Create,
                changes.ItemsElementName[0],
                "After updating the item, if the SyncState element is not specified when calling SyncFolderItems, the changes between items on the client and the items on the server should be 'Create'");
            Site.Assert.AreEqual<string>(
                newItemSubject,
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.Subject,
                "After updating the item, the subject of the item should be the expected one.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R382");

            // If the changes is Create and the subject of the item is updated, it indicates the item in its current state is returned as if it has never been synchronized,
            // then requirement MS-OXWSSYNC_R382 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R382
            Site.CaptureRequirement(
                382,
                @"[In m:SyncFolderItemsType Complex Type] If this element [SyncState] is not specified, all synchronization information is returned.");
            #endregion

            #region Step 6. Client invokes DeleteItem operation to delete the ContactItemType item which updated in Step 4.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ContactItemType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ContactItemType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync PostItemType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC07_SyncFolderItems_PostItemType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of inbox folder and verify related requirements.
            DistinguishedFolderIdNameType inboxFolder = DistinguishedFolderIdNameType.inbox;

            // Call SyncFolderItems operation with invalid SyncState to verify the error code: ErrorInvalidSyncStateData
            SyncFolderItemsType requestWithInvalidSyncState = this.CreateSyncFolderItemsRequestWithoutOptionalElements(inboxFolder, DefaultShapeNamesType.AllProperties);

            // The SyncState element data, encoded with base64 encoding, is set to an invalid value
            requestWithInvalidSyncState.SyncState = TestSuiteBase.InvalidSyncState;
            SyncFolderItemsResponseType responseWithInvalidSyncState = this.SYNCAdapter.SyncFolderItems(requestWithInvalidSyncState);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R518");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R518
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidSyncStateData,
                responseWithInvalidSyncState.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                518,
                @"[In m:ResponseCodeType Simple Type] [ErrorInvalidSyncStateData: ] This is returned by the SyncFolderItems method if the SyncState property data is invalid.");

            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(inboxFolder, DefaultShapeNamesType.AllProperties);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem create a PostItemType item and get its ID.
            PostItemType postItem = new PostItemType();
            BaseItemIdType[] itemIds = this.CreateItem(inboxFolder, postItem);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one PostItemType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is PostItemType, then requirement MS-OXWSSYNC_R174 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R174");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R174
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(PostItemType),
                174,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of PostItem is t:PostItemType ([MS-OXWSPOST] section 2.2.4.1).");

            // Assert both the length of responseMessage.Changes.ItemsElementName and responseMessage.Changes.Items are 1.
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one PostItemType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isPostItemCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                        (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(PostItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1751. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(PostItemType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is PostItemType, it indicates a post item 
            // has been created on server and synced on client, then requirement MS-OXWSSYNC_R1751 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1751
            Site.CaptureRequirementIfIsTrue(
                isPostItemCreated,
                1751,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element PostItem] specifies a post item to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, inboxFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one PostItemType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            bool isLastItemIncluded = responseMessage.IncludesLastItemInRange &&
               ((changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(PostItemType));

            // Since the last updated item is a post item, if the IncludesLastItemInRange element in SyncFolderItems response is TRUE 
            // and the items in Changes contains PostItemType item, then requirement MS-OXWSSYNC_R6701 can be captured.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R6701. Expected value: IncludesLastItemInRange: 'true', item type: {0}; actual value: IncludesLastItemInRange: {1}, item type: {2}",
                typeof(PostItemType),
                responseMessage.IncludesLastItemInRange,
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R6701
            Site.CaptureRequirementIfIsTrue(
                isLastItemIncluded,
                6701,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] [The element IncludesLastItemInRange] True indicates the last item to synchronize is included in the response.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one PostItemType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isPostItemUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(PostItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1752. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(PostItemType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is PostItemType, it indicates a post item 
            // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1752 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1752
            Site.CaptureRequirementIfIsTrue(
                isPostItemUpdated,
                1752,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element PostItem] specifies a post item to update in the client data store.");
            #endregion

            #region Step 6. Client invokes DeleteItem operation to delete the PostItemType item which updated in Step 4.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one PostItemType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one PostItemType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync CalendarItemType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC08_SyncFolderItems_CalendarItemType()
        {
            #region Step 1. Client invokes SyncFolderItems operation get initial syncState of calendar folder.
            DistinguishedFolderIdNameType calendarFolder = DistinguishedFolderIdNameType.calendar;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(calendarFolder, DefaultShapeNamesType.IdOnly);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem create a CalendarItemType item and get its ID.
            CalendarItemType calendarItem = new CalendarItemType();
            BaseItemIdType[] itemIds = this.CreateItem(calendarFolder, calendarItem);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            SyncFolderItemsType requestWithNormalItems = this.CreateSyncFolderItemsRequestWithoutOptionalElements(calendarFolder, DefaultShapeNamesType.AllProperties);

            // Include SyncState element and set its value to the one that got from the first synchronization
            requestWithNormalItems.SyncState = responseMessage.SyncState;
            if (Common.IsRequirementEnabled(37811008, this.Site))
            {
                // Set the value of SyncScope to "NormalItems"
                requestWithNormalItems.SyncScopeSpecified = true;
                requestWithNormalItems.SyncScope = SyncFolderItemsScopeType.NormalItems;
            }

            // Set the value of AdditionalProperties element
            requestWithNormalItems.ItemShape.AdditionalProperties = new BasePathToElementType[] 
            { 
                new PathToUnindexedFieldType() 
                { 
                    FieldURI = UnindexedFieldURIType.itemSubject
                } 
            };

            SyncFolderItemsResponseType responseWithNormalItems = this.SYNCAdapter.SyncFolderItems(requestWithNormalItems);
            SyncFolderItemsResponseMessageType responseMessageWithNormalItems = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(responseWithNormalItems);
            if (Common.IsRequirementEnabled(37811008, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "MS-OXWSSYNC_R37811008");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811008
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    responseMessage.ResponseCode,
                    37811008,
                    @"[In Appendix C: Product Behavior] Implementation dose support the SyncScope element. (Exchange 2010 and above follow this behavior.)");
            }

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changesWithNormalItems = responseMessageWithNormalItems.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changesWithNormalItems.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changesWithNormalItems.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changesWithNormalItems.Items.Length, "Just one CalendarItemType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is CalendarItemType, then requirement MS-OXWSSYNC_R158 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R158");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R158
            Site.CaptureRequirementIfIsInstanceOfType(
                (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(CalendarItemType),
                158,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of CalendarItem is t:CalendarItemType ([MS-OXWSMTGS] section 2.2.4.4).");

            // If the value of SyncScope is set to 'NormalItems', there should be only items in the folder returned and the value of IsAssociated property of the items in the folder should be false
            Site.Assert.IsFalse(
                (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.IsAssociated,
                "The folder associated items should not be returned if the value of SyncScope is set to 'NormalItems'.");
            if (Common.IsRequirementEnabled(37811008, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "MS-OXWSSYNC_R347");

                // If there is only a CalendarType item in Changes and the IsAssociated property is false, it indicates if the SyncScope is set to "NormalItems", only the item in the folder is returned.
                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R347
                Site.CaptureRequirementIfIsInstanceOfType(
                    (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                    typeof(CalendarItemType),
                    347,
                    @"[In t:SyncFolderItemsScopeType Simple Type] [The value NormalItems] specifies that only items in the folder are returned.");
            }

            // If the AdditionalProperties element is included in SyncFolderItems request and the FieldURI is point to item subject, 
            // the additional property subject should be returned in response, then requirement MS-OXWSSYNC_R37814 and MS-OXWSSYNC_R37815 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R37814");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37814
            Site.CaptureRequirementIfIsNotNull(
                (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.Subject,
                37814,
                @"[In m:SyncFolderItemsType Complex Type] [ItemShape element AdditionalProperties] Specifies a set of requested additional properties to return in a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R37815");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37815
            Site.CaptureRequirementIfIsNotNull(
                (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.Subject,
                37815,
                @"[In m:SyncFolderItemsType Complex Type] [ItemShape element AdditionalProperties, element t:Path] Specifies a property to be returned in a response.");

            // Assert both the length of responseMessage.Changes.ItemsElementName and responseMessage.Changes.Items are 1.
            Site.Assert.AreEqual<int>(1, changesWithNormalItems.ItemsElementName.Length, "Just one CalendarItemType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isCalendarItemCreated = changesWithNormalItems.ItemsElementName[0] == ItemsChoiceType1.Create &&
                        (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(CalendarItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1591. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(CalendarItemType),
                changesWithNormalItems.ItemsElementName[0],
                (changesWithNormalItems.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is CalendarItemType, it indicates a calendar item 
            // has been created on server and synced on client, then requirement MS-OXWSSYNC_R1591 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1591
            Site.CaptureRequirementIfIsTrue(
                isCalendarItemCreated,
                1591,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element CalendarItem] specifies a calendar item to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, calendarFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(calendarFolder, responseMessageWithNormalItems, DefaultShapeNamesType.AllProperties);
            if (Common.IsRequirementEnabled(37811008, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "MS-OXWSSYNC_R37811008");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811008
                Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    responseMessage.ResponseCode,
                    37811008,
                    @"[In Appendix C: Product Behavior] Implementation dose support the SyncScope element. (Exchange 2010 and above follow this behavior.)");
            }

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one CalendarItemType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one CalendarItemType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isCalendarItemUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(CalendarItemType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1592. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(CalendarItemType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is CalendarItemType, it indicates a calendar item 
            // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1592 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1592
            Site.CaptureRequirementIfIsTrue(
                isCalendarItemUpdated,
                1592,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type][The element CalendarItem] specifies a calendar item to update in the client data store.");
            #endregion

            #region Step 6. Client invokes DeleteItem operation to delete the CalendarItemType item which updated in Step 4.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(calendarFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one CalendarItemType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one CalendarItemType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync DistributionListType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC09_SyncFolderItems_DistributionListType()
        {
            // Check whether the DistributionListType item is supported on current server version
            Site.Assume.IsTrue(Common.IsRequirementEnabled(37811, this.Site), "Exchange 2007 does not support DistributionListType item, for detailed information refer to MS-OXWSDLIST.");

            #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of contacts folder.
            DistinguishedFolderIdNameType contactFolder = DistinguishedFolderIdNameType.contacts;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(contactFolder, DefaultShapeNamesType.AllProperties);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem to create a DistributionListType item and get its ID.
            DistributionListType distributionList = new DistributionListType();
            BaseItemIdType[] itemIds = this.CreateItem(contactFolder, distributionList);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2 and verify related requirements.
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one DistributionList item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If the type of item in SyncFolderItems response is DistributionListType, then requirement MS-OXWSSYNC_R37811 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R37811");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811
            Site.CaptureRequirementIfIsInstanceOfType(
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                typeof(DistributionListType),
                37811,
                @"[In Appendix C: Product Behavior] Implementation does support DistributionList with type t:DistributionListType ([MS-OXWSDLIST] section 2.2.4.3). (Exchange 2010 and above follow this behavior.)");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one DistributionListType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            bool isDistributionListCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                        (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(DistributionListType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1631. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Create,
                typeof(DistributionListType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Create and the type of Item is DistributionListType, it indicates a distribution list 
            // has been created on server and synced on client, then requirement MS-OXWSSYNC_R1631 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1631
            Site.CaptureRequirementIfIsTrue(
                isDistributionListCreated,
                1631,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element DistributionList] specifies a distribution list to create in the client data store.");
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, contactFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one DistributionList item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one DistributionList item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            bool isDistributionListUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(DistributionListType);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXWSSYNC_R1632. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                ItemsChoiceType1.Update,
                typeof(DistributionListType),
                changes.ItemsElementName[0],
                (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

            // If the ItemsElementName of Changes is Update and the type of Item is DistributionListType, it indicates a distribution list 
            // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R1592 can be captured.
            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1632
            Site.CaptureRequirementIfIsTrue(
                isDistributionListUpdated,
                1632,
                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] [The element DistributionList] specifies a distribution list to update in the client data store.");
            #endregion

            #region Step 6. Client invokes DeleteItem operation to delete the DistributionListType item which updated in Step 4.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            // Assert both the length of responseMessage.Changes.ItemsElementName and responseMessage.Changes.Items are 1.
            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one DistributionList item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

            // Assert the ItemsElementName is Delete.
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one DistributionList item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // Assert the Items is an instance of SyncFolderItemsDeleteType.
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync ItemType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC10_SyncFolderItems_ItemType()
        {
            #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of inbox folder.
            DistinguishedFolderIdNameType inboxFolder = DistinguishedFolderIdNameType.inbox;
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(inboxFolder, DefaultShapeNamesType.AllProperties);
            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            #endregion

            #region Step 2. Client invokes CreateItem to create a ItemType item and get its ID.
            ItemType item = new ItemType();
            BaseItemIdType[] itemIds = this.CreateItem(inboxFolder, item);
            #endregion

            #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
            SyncFolderItemsChangesType changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ItemType item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Create,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Create', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ItemType item was created in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If client creates an item of ItemType, a MessageType complex type is returned.
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsCreateOrUpdateType) && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MessageType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}' and the type of Item should be '{1}'.", typeof(SyncFolderItemsCreateOrUpdateType), typeof(MessageType)));
            if (Common.IsRequirementEnabled(37811004, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWWSSYNC_R37811004");

                this.Site.CaptureRequirementIfIsInstanceOfType(
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                    typeof(MessageType),
                    37811004,
                    @"[In Appendix C: Product Behavior] Implementation dose return a MessageType complex type. (If a client creates an item of this type, a MessageType complex type is returned.)");
            }
            #endregion

            #region Step 4. Client invokes UpdateItem operation to update the created item which created in Step 2.
            // Generate a new item subject
            string newItemSubject = Common.GenerateResourceName(this.Site, inboxFolder + "NewItemSubject");
            this.UpdateItemSubject(itemIds, newItemSubject);
            #endregion

            #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ItemType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Update,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Update', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ItemType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");

            // If client creates an item of ItemType, a MessageType complex type is returned.
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsCreateOrUpdateType) && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(MessageType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}' and the type of Item should be '{1}'.", typeof(SyncFolderItemsCreateOrUpdateType), typeof(MessageType)));
            #endregion

            #region Step 6. Client invokes DeleteItem operation to delete the ItemType item which updated in Step 4.
            this.DeleteItem(itemIds);
            #endregion

            #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
            responseMessage = this.GetResponseMessage(inboxFolder, responseMessage, DefaultShapeNamesType.AllProperties);

            // Assert the changes in response is not null
            Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
            changes = responseMessage.Changes;

            // Assert both the Items and ItemsElementName are not null
            Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
            Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

            Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ItemType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

            Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one ItemType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");
            Site.Assert.IsTrue(
                changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));
            #endregion
        }

        /// <summary>
        /// Client calls SyncFolderItems operation to sync ItemType item.
        /// </summary>
        [TestCategory("MSOXWSSYNC"), TestMethod()]
        public void MSOXWSSYNC_S02_TC11_SyncFolderItems_AbchPersonItemType()
        {
                Site.Assume.IsTrue(Common.IsRequirementEnabled(37811006, this.Site), "Implementation dose support the Person element.");
           
                #region Step 1. Client invokes SyncFolderItems operation to get initial syncState of contacts folder.
                DistinguishedFolderIdNameType contactFolder = DistinguishedFolderIdNameType.contacts;
                SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(contactFolder, DefaultShapeNamesType.AllProperties);
                SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
                SyncFolderItemsResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
                #endregion

                #region Step 2. Client invokes CreateItem to create a DistributionListType item and get its ID.
                AbchPersonItemType distributionList = new AbchPersonItemType();
                BaseItemIdType[] itemIds = this.CreateItem(contactFolder, distributionList);
                #endregion

                #region Step 3. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 2.
                responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

                // Assert the changes in response is not null
                Site.Assert.IsNotNull(responseMessage.Changes, "There is one item created on server, so the changes between server and client should not be null");
                SyncFolderItemsChangesType changes = responseMessage.Changes;

                // Assert both the Items and ItemsElementName are not null
                Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item created on server.");
                Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item created on server.");
                Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one ABchPersonItemtype item was created in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
                
                // If the type of item in SyncFolderItems response is AbchPersonItemType, then requirement MS-OXWSSYNC_R1752005 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R1752005");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R1752005
                Site.CaptureRequirementIfIsInstanceOfType(
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item,
                    typeof(AbchPersonItemType),
                    1752005,
                    @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The type of  Person is t:AbchPersonItemType ([MS-OXWSCONT] section 2.2.4.1)");

                bool isDistributionListCreated = changes.ItemsElementName[0] == ItemsChoiceType1.Create &&
                            (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(AbchPersonItemType);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXWSSYNC_R37811006. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                    ItemsChoiceType1.Create,
                    typeof(AbchPersonItemType),
                    changes.ItemsElementName[0],
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

                // If the ItemsElementName of Changes is Create and the type of Item is AbchPersonItemType, it indicates a AbchPersonItemType list 
                // has been created on server and synced on client, then requirement MS-OXWSSYNC_R37811006 can be captured.
                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811006
                Site.CaptureRequirementIfIsTrue(
                    isDistributionListCreated,
                    37811006,
                    @"[In Appendix C: Product Behavior] Implementation dose support the Person element. (Exchange 2016 follow this behavior.)");
                #endregion               

                #region Step 4. Client invokes UpdateItem operation to update the created item which created in Step 2.
                // Generate a new item subject
                string newItemSubject = Common.GenerateResourceName(this.Site, contactFolder + "NewItemSubject");
                this.UpdateItemSubject(itemIds, newItemSubject);
                #endregion

                #region Step 5. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 4 and verify related requirements.
                responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

                // Assert the changes in response is not null
                Site.Assert.IsNotNull(responseMessage.Changes, "There is one item updated on server, so the changes between server and client should not be null");
                changes = responseMessage.Changes;

                // Assert both the Items and ItemsElementName are not null
                Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item updated on server.");
                Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item updated on server.");

                Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one AbchPersonItemType item was updated in previous step, so the count of Items array in responseMessage.Changes should be 1.");
                Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one AbchPersonItemType item was updated in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");
                
                bool isDistributionListUpdated = changes.ItemsElementName[0] == ItemsChoiceType1.Update
                    && (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType() == typeof(AbchPersonItemType);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXWSSYNC_R37811006. Expected value: ItemsElementName: {0}, item type: {1}; actual value: ItemsElementName: {2}, item type: {3}",
                    ItemsChoiceType1.Update,
                    typeof(AbchPersonItemType),
                    changes.ItemsElementName[0],
                    (changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.GetType());

                // If the ItemsElementName of Changes is Update and the type of Item is DistributionListType, it indicates a distribution list 
                // has been updated on server and synced on client, then requirement MS-OXWSSYNC_R37811006 can be captured.
                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R37811006
                Site.CaptureRequirementIfIsTrue(
                    isDistributionListUpdated,
                    37811006,
                    @"[In Appendix C: Product Behavior] Implementation dose support the Person element. (Exchange 2016 follow this behavior.)");
                #endregion
                
                #region Step 6. Client invokes DeleteItem operation to delete the DistributionListType item which updated in Step 4.
                this.DeleteItem(itemIds);
                #endregion

                #region Step 7. Client invokes SyncFolderItems operation with previous SyncState to sync the operation result in Step 6.
                responseMessage = this.GetResponseMessage(contactFolder, responseMessage, DefaultShapeNamesType.AllProperties);

                // Assert the changes in response is not null
                Site.Assert.IsNotNull(responseMessage.Changes, "There is one item deleted on server, so the changes between server and client should not be null");
                changes = responseMessage.Changes;

                // Assert both the Items and ItemsElementName are not null
                Site.Assert.IsNotNull(changes.ItemsElementName, "There should be changes information returned in SyncFolderItems response since there is one item deleted on server.");
                Site.Assert.IsNotNull(changes.Items, "There should be item information returned in SyncFolderItems response since there is one item deleted on server.");

                // Assert both the length of responseMessage.Changes.ItemsElementName and responseMessage.Changes.Items are 1.
                Site.Assert.AreEqual<int>(1, changes.ItemsElementName.Length, "Just one AbchPersonItemType item was deleted in previous step, so the count of ItemsElementName array in responseMessage.Changes should be 1.");

                // Assert the ItemsElementName is Delete.
                Site.Assert.IsTrue(
                    changes.ItemsElementName[0] == ItemsChoiceType1.Delete,
                    string.Format("The responseMessage.Changes.ItemsElementName should be 'Delete', the actual value is '{0}'", changes.ItemsElementName[0]));

                Site.Assert.AreEqual<int>(1, changes.Items.Length, "Just one AbchPersonItemType item was deleted in previous step, so the count of Items array in responseMessage.Changes should be 1.");

                // Assert the Items is an instance of SyncFolderItemsDeleteType.
                Site.Assert.IsTrue(
                    changes.Items[0].GetType() == typeof(SyncFolderItemsDeleteType),
                    string.Format("The responseMessage.Changes.Items should be an instance of '{0}'.", typeof(SyncFolderItemsDeleteType)));
                #endregion
            
        }
        #endregion

    } 
}