namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The bass class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// An invalid value for SyncState element which is encoded with base64 encoding
        /// </summary>
        protected const string InvalidSyncState = "H4sIAAA==";

        /// <summary>
        /// The search text of a search folder
        /// </summary>
        protected const string SearchText = "Search Text";

        #endregion

        #region Properties
        /// <summary>
        /// Gets the email address of User2.
        /// </summary>
        protected string User2EmailAddress { get; private set; }

        /// <summary>
        /// Gets the list which stores DistinguishedFolderIdNameType type folders.
        /// </summary>
        protected List<DistinguishedFolderIdNameType> FolderIdNameType { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSSYNC protocol adapter instance.
        /// </summary>
        protected IMS_OXWSSYNCAdapter SYNCAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSCORE protocol adapter instance.
        /// </summary>
        protected IMS_OXWSCOREAdapter COREAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSSYNC SUT control adapter instance.
        /// </summary>
        protected IMS_OXWSSYNCSUTControlAdapter SYNCSUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSFOLD SUT control adapter instance.
        /// </summary>
        protected IMS_OXWSFOLDSUTControlAdapter FOLDSUTControlAdapter { get; private set; }
        #endregion

        #region Test case initialize and clean up
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.SYNCAdapter = Site.GetAdapter<IMS_OXWSSYNCAdapter>();
            this.COREAdapter = Site.GetAdapter<IMS_OXWSCOREAdapter>();
            this.SYNCSUTControlAdapter = Site.GetAdapter<IMS_OXWSSYNCSUTControlAdapter>();
            this.FOLDSUTControlAdapter = Site.GetAdapter<IMS_OXWSFOLDSUTControlAdapter>();

            // Get the email address value of User2.
            this.User2EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            this.FolderIdNameType = new List<DistinguishedFolderIdNameType> 
            {
                DistinguishedFolderIdNameType.calendar,
                DistinguishedFolderIdNameType.contacts,
                DistinguishedFolderIdNameType.tasks
            };
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // Clean up the mailbox of User1.
            bool isCleaned = this.SYNCSUTControlAdapter.CleanupMailBox(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site));

            Site.Assert.IsTrue(
                isCleaned,
                string.Format("All the items and sub folders in the mailbox of user '{0}' should be cleaned.", Common.GetConfigurationPropertyValue("User1Name", this.Site)));
            base.TestCleanup();
        }
        #endregion

        #region Test case base methods
        #region Test case helper
        /// <summary>
        /// Clean up the mailbox of User2.
        /// </summary>
        protected void CleanupAttendeeMailbox()
        {
            // Clean up the mailbox of User2.
            bool isCleaned = this.SYNCSUTControlAdapter.CleanupMailBox(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site));

            Site.Assert.IsTrue(
                isCleaned,
                string.Format("All the items and sub folders in mailbox of user '{0}' should be cleaned.", Common.GetConfigurationPropertyValue("User2Name", this.Site)));
        }

        /// <summary>
        /// Creates and sends a meeting request message.
        /// </summary>
        /// <param name="destinationUserEmail">Email address of the person to whom meeting request which should be sent to.</param>
        /// <param name="itemSubject">Subject of the meeting request which should be sent.</param>
        /// <returns>The ID of the meeting request.</returns>
        protected ItemIdType CreateMeetingRequest(string destinationUserEmail, string itemSubject)
        {
            // Create a request for the CreateItem operation.
            CreateItemType createItemRequest = new CreateItemType();

            // Add the CalendarItemType item to the items to be created.
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();

            // Create a CalendarItemType item.
            CalendarItemType calendarItem = new CalendarItemType();

            // Set the receiver.
            calendarItem.RequiredAttendees = new AttendeeType[1];
            calendarItem.RequiredAttendees[0] = new AttendeeType();
            calendarItem.RequiredAttendees[0].Mailbox = new EmailAddressType();
            calendarItem.RequiredAttendees[0].Mailbox.EmailAddress = destinationUserEmail;
            calendarItem.Subject = itemSubject;

            // Add the calendar item to the request.
            createItemRequest.Items.Items = new ItemType[] { calendarItem };

            // Set the SendMeetingInvitations property, send to all and save copy.
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
            createItemRequest.SendMeetingInvitationsSpecified = true;

            ItemIdType itemId = null;
            itemId = this.CreateMeetingMessage(createItemRequest);

            #region Check the sent items folder and calendar folder of User1
            // Make sure that the meeting request message has been saved to sent items folder of User1.
            bool isSavedToSentItems = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.sentitems.ToString(),
                itemSubject,
                Item.MeetingRequest.ToString());
            Site.Assert.IsTrue(
                isSavedToSentItems,
                string.Format("The meeting request message should be saved to sent items folder of '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site)));

            // Make sure that the meeting request item exists in calendar folder of User1.
            bool isCreateInCalendar = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.calendar.ToString(),
                itemSubject,
                Item.CalendarItem.ToString());
            Site.Assert.IsTrue(
                isCreateInCalendar,
                string.Format("The meeting request item should exist in calendar folder of '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site)));
            #endregion

            #region Check the calendar folder of User2
            // Make sure that the meeting request item exists in calendar folder of User2.
            bool isRequestItemInCalendar = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.calendar.ToString(),
                itemSubject,
                Item.CalendarItem.ToString());
            Site.Assert.IsTrue(
                isRequestItemInCalendar,
                string.Format("The meeting request item should exist in calendar folder of '{0}'", Common.GetConfigurationPropertyValue("User2Name", this.Site)));
            #endregion

            return itemId;
        }

        /// <summary>
        /// Creates a new meeting response message.
        /// </summary>
        /// <param name="itemSubject">Subject of the item which should be created.</param>
        protected void CreateMeetingResponse(string itemSubject)
        {
            // User1 sends a meeting request to User2.
            ItemIdType sendItemId = this.CreateMeetingRequest(this.User2EmailAddress, itemSubject);
            Site.Assert.IsNotNull(
                sendItemId,
                string.Format("The meeting request should be sent to '{0}' successfully.", this.User2EmailAddress));

            // User2 accepts the meeting request message in the inbox folder.
            bool isAccepted = this.SYNCSUTControlAdapter.FindAndAcceptMeetingMessage(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                itemSubject,
                Item.MeetingRequest.ToString());
            Site.Assert.IsTrue(
                isAccepted,
                string.Format(
                "The User2 '{0}' should accept the meeting request message named '{1}'.", Common.GetConfigurationPropertyValue("User2Name", this.Site), itemSubject));

            #region Check the inbox folder of User1
            // Make sure that the meeting response message is received by User1.
            bool isReceived = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.inbox.ToString(),
                itemSubject,
                Item.MeetingResponse.ToString());
            Site.Assert.IsTrue(
                isReceived,
                string.Format("The meeting acceptation message should be received by '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site)));
            #endregion

            #region Check deleted items folder of User2
            // Make sure that the meeting request message is moved to deleted items folder of User2.
            bool isRequestDeleted = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.deleteditems.ToString(),
                itemSubject,
                Item.MeetingRequest.ToString());
            Site.Assert.IsTrue(isRequestDeleted, "The meeting request message should be moved to deleted items folder of '{0}'", Common.GetConfigurationPropertyValue("User2Name", this.Site));
            #endregion
        }

        /// <summary>
        /// Creates a new meeting cancellation message.
        /// </summary>
        /// <param name="itemSubject">Subject of the item which should be created.</param>
        protected void CreateMeetingCancellation(string itemSubject)
        {
            // User1 sends a meeting request to User2.
            ItemIdType meetingRequestCalendar = this.CreateMeetingRequest(this.User2EmailAddress, itemSubject);
            Site.Assert.IsNotNull(meetingRequestCalendar, "The meeting request to '{0}' should be sent successfully.", this.User2EmailAddress);

            // Make sure that user2 received the meeting request.
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

            // User1 cancels the meeting and saves a meeting cancellation message copy into its inbox folder.
            bool cancelMeeting = this.CancelMeeting(meetingRequestCalendar);
            Site.Assert.IsTrue(cancelMeeting, "The meeting named '{0}' should be cancelled successfully.", itemSubject);

            #region Check the junkemail folder and deleted items folder of User1
            // Make sure that the meeting cancellation message exists in User1's junkemail folder.
            bool isCancellationReceived = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.junkemail.ToString(),
                itemSubject,
                Item.MeetingCancellation.ToString());
            Site.Assert.IsTrue(isCancellationReceived, "The meeting cancellation message should be received by '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site));

            // Make sure that the meeting request item exists in User1's deleted items folder.
            bool isRequestItem = this.SYNCSUTControlAdapter.IsItemExisting(
               Common.GetConfigurationPropertyValue("User1Name", this.Site),
               Common.GetConfigurationPropertyValue("User1Password", this.Site),
               Common.GetConfigurationPropertyValue("Domain", this.Site),
               DistinguishedFolderIdNameType.deleteditems.ToString(),
               itemSubject,
               Item.CalendarItem.ToString());
            Site.Assert.IsTrue(isRequestItem, "The meeting request message should exist in deleted items folder of '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site));
            #endregion

            #region Check the inbox folder, deleted items folder and calendar folder of User2
            // Make sure that the meeting cancellation message exists in User2's inbox folder
            isCancellationReceived = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.inbox.ToString(),
                itemSubject,
                Item.MeetingCancellation.ToString());
            Site.Assert.IsTrue(isCancellationReceived, "The meeting cancellation message should be received by '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site));

            // Make sure that the meeting request message exists in User2's inbox folder or deleted items folder
            bool isRequestMessageInDeletedItems = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.deleteditems.ToString(),
                itemSubject,
                Item.MeetingRequest.ToString());
            bool isRequestMessageInInbox = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.inbox.ToString(),
                itemSubject,
                Item.MeetingRequest.ToString());
            bool isRequestMessageDeleted = isRequestMessageInDeletedItems || isRequestMessageInInbox;
            Site.Assert.IsTrue(isRequestMessageDeleted, "The meeting request message should exist in inbox or deleted items folder of '{0}'", Common.GetConfigurationPropertyValue("User1Name", this.Site));

            // Make sure that the meeting cancellation item exists in User2's calendar folder
            isCancellationReceived = this.SYNCSUTControlAdapter.IsItemExisting(
                Common.GetConfigurationPropertyValue("User2Name", this.Site),
                Common.GetConfigurationPropertyValue("User2Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                DistinguishedFolderIdNameType.calendar.ToString(),
                itemSubject,
                Item.CalendarItem.ToString());
            Site.Assert.IsTrue(isCancellationReceived, "The meeting cancellation item should exist in calendar folder of '{0}'", Common.GetConfigurationPropertyValue("User2Name", this.Site));
            #endregion
        }
        #endregion

        #region Construct MS-OXWSSYNC protocol adapter method request
        /// <summary>
        /// Creates the SyncFolderItems request without optional elements.
        /// </summary>
        /// <param name="folderIdName">A default folder name.</param>
        /// <param name="defaultShapeNamesType">Standard sets of properties to return.</param>
        /// <returns>A SyncFolderItemsType request.</returns>
        protected SyncFolderItemsType CreateSyncFolderItemsRequestWithoutOptionalElements(
            DistinguishedFolderIdNameType folderIdName,
            DefaultShapeNamesType defaultShapeNamesType)
        {
            // Get the request of CreateSyncFolderItems operation.
            SyncFolderItemsType request = new SyncFolderItemsType();
            request.SyncFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = folderIdName;
            request.SyncFolderId.Item = distinguishedFolderId;
            request.MaxChangesReturned = int.Parse(Common.GetConfigurationPropertyValue("MaxChanges", this.Site));

            request.ItemShape = new ItemResponseShapeType();
            request.ItemShape.BaseShape = defaultShapeNamesType;

            return request;
        }

        /// <summary>
        /// Get SyncFolderHierarchy response, the request doesn't include all optional elements.
        /// </summary>
        /// <returns>A SyncFolderHierarchyResponseMessageType response message.</returns>
        protected SyncFolderHierarchyResponseMessageType GetSyncFolderHierarchyResponseMessage()
        {
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest();

            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            return responseMessage;
        }

        /// <summary>
        /// Get SyncFolderHierarchy response, the request does include all optional elements.
        /// </summary>
        /// <param name="responseMessages">A response message of SyncFolderHierarchyResponseMessageType.</param>
        /// <param name="folder">A default folder name.</param>
        /// <returns>A SyncFolderHierarchyResponseMessageType response message.</returns>
        protected SyncFolderHierarchyResponseMessageType GetSyncFolderHierarchyResponseMessage(
            SyncFolderHierarchyResponseMessageType responseMessages,
            DistinguishedFolderIdNameType folder)
        {
            SyncFolderHierarchyType request = TestSuiteHelper.CreateSyncFolderHierarchyRequest(folder, DefaultShapeNamesType.Default, true, true);
            request.SyncState = responseMessages.SyncState;
            SyncFolderHierarchyResponseType response = this.SYNCAdapter.SyncFolderHierarchy(request);
            SyncFolderHierarchyResponseMessageType responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderHierarchyResponseMessageType>(response);

            return responseMessage;
        }
        #endregion

        #region Construct MS-OXWSCORE protocol adapter method request and get response
        /// <summary>
        /// Creates an item under a specified folder.
        /// </summary>
        /// <typeparam name="T">Type of the item which should be created.</typeparam>
        /// <param name="folderId">Parent folder of the item which should be created.</param>
        /// <param name="item">An instance of ItemType.</param>
        /// <returns>The array of item IDs of the created item.</returns>
        protected ItemIdType[] CreateItem<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType
        {
            CreateItemType request = new CreateItemType();

            // Set the sendMeetingInvitations property.
            request.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            request.SendMeetingInvitationsSpecified = true;

            // Set the MessageDisposition property.
            request.MessageDisposition = MessageDispositionType.SaveOnly;
            request.MessageDispositionSpecified = true;

            // Specify the folder in which the new items should be saved.
            request.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType parentFolderId = new DistinguishedFolderIdType();
            parentFolderId.Id = folderId;
            request.SavedItemFolderId.Item = parentFolderId;
            item.Subject = Common.GenerateResourceName(this.Site, folderId + "ItemSubject");

            // Specify the collection of items to be created.
            request.Items = new NonEmptyArrayOfAllItemsType();
            request.Items.Items = new T[] { item };

            // Invoke the CreateItem operation.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(request);

            // Check whether the CreateItem operation is executed successfully.
            Site.Assert.AreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    createItemResponse.ResponseMessages.Items[0].ResponseClass,
                    string.Format(
                        "Create item should be successful! Expected response code: {0}, actual response code: {1}",
                        ResponseCodeType.NoError,
                        createItemResponse.ResponseMessages.Items[0].ResponseCode));

            // Get the item ID.
            ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(1, itemIds.Length, "There should be 1 item returned!");

            return itemIds;
        }

        /// <summary>
        /// Cancels the specified meeting request.
        /// </summary>
        /// <param name="itemId">ID of the item which should be cancelled.</param>
        /// <returns>If the meeting request is cancelled successfully, return true; otherwise, return false.</returns>
        protected bool CancelMeeting(ItemIdType itemId)
        {
            // Create a request for the CreateItem operation.
            CreateItemType createItemRequest = new CreateItemType();

            // Add the CalendarItemType item to the items to be created.
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();

            // Create a CancelCalendarItemType item to cancel a meeting
            CancelCalendarItemType cancelItem = new CancelCalendarItemType();

            // Set the related meeting calendar item.
            cancelItem.ReferenceItemId = new ItemIdType();
            cancelItem.ReferenceItemId = itemId;
            createItemRequest.Items.Items = new ItemType[] { cancelItem };

            // Set the saved folder to the junkemail folder.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType folder;
            folder = new DistinguishedFolderIdType();
            folder.Id = DistinguishedFolderIdNameType.junkemail;
            createItemRequest.SavedItemFolderId.Item = folder;

            // Set the MessageDisposition property to SendAndSaveCopy.
            createItemRequest.MessageDisposition = MessageDispositionType.SendAndSaveCopy;
            createItemRequest.MessageDispositionSpecified = true;
            bool isCancelled = false;

            // If the meeting cancellation message is created successfully, return true; otherwise, return false.
            if (this.CreateMeetingMessage(createItemRequest) != null)
            {
                isCancelled = true;           
            }

            return isCancelled;
        }

        /// <summary>
        /// Gets the IsRead property value of specified items.
        /// </summary>
        /// <param name="itemIds">ID of the items to get.</param>
        /// <returns>Value of the IsRead property.</returns>
        protected string[] GetReadFlag(BaseItemIdType[] itemIds)
        {
            // Create a request for the GetItem operation.
            GetItemType request = new GetItemType();

            // Specify the item to be gotten.
            request.ItemIds = itemIds;

            // Return all the properties that are defined for the AllProperties shape.
            request.ItemShape = new ItemResponseShapeType();
            request.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // Invoke the GetItem operation.
            GetItemResponseType response = this.COREAdapter.GetItem(request);

            // Check whether the GetItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Get item should be successful! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }

            // Get the response message.
            ItemInfoResponseMessageType info = response.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            string[] isRead = new string[info.Items.Items.Length];
            for (int i = 0; i < info.Items.Items.Length; i++)
            {
                // Find the IsRead property of the item and get the value.
                Type type = info.Items.Items[i].GetType();
                PropertyInfo isReadFlagSpecified = type.GetProperty("IsReadSpecified");

                if ((bool)isReadFlagSpecified.GetValue(info.Items.Items[i], null))
                {
                    PropertyInfo readFlag = type.GetProperty("IsRead");
                    isRead[i] = readFlag.GetValue(info.Items.Items[i], null).ToString();
                }
            }

            return isRead;
        }

        /// <summary>
        /// Marks read flag of the specified item, depending on the parameter isRead.
        /// </summary>
        /// <param name="itemIds">ID of the item which should be updated.</param>
        /// <param name="isRead">The read flag current status.</param>
        protected void UpdateReadFlag(BaseItemIdType[] itemIds, bool[] isRead)
        {
            ItemChangeType[] itemChanges = new ItemChangeType[itemIds.Length];

            // Set the public properties (Subject) which all the seven kinds of operation have.
            for (int i = 0; i < itemIds.Length; i++)
            {
                itemChanges[i] = new ItemChangeType();
                itemChanges[i].Item = itemIds[i];
                itemChanges[i].Updates = new ItemChangeDescriptionType[]
                    {
                        new SetItemFieldType()
                        {
                            Item = new PathToUnindexedFieldType()
                            {
                                FieldURI = UnindexedFieldURIType.messageIsRead
                            },
                            Item1 = new MessageType()
                            {
                               IsReadSpecified = true,
                               IsRead = isRead[i]
                            }
                        }
                    };
            }

            // Update the item.
            this.UpdateItem(itemChanges);
        }

        /// <summary>
        /// Convert the read flag of items
        /// </summary>
        /// <param name="itemIds">Id of the items which should be converted.</param>
        /// <returns>The converted read flag.</returns>
        protected bool[] ConvertReadFlag(BaseItemIdType[] itemIds)
        {
            // Call GetReadFlag to get the IsRead property.
            string[] readFlag = this.GetReadFlag(itemIds);

            // An array to store the Boolean value of isRead 
            bool[] isRead = new bool[readFlag.Length];

            // An array to store the opposite value of isRead.
            bool[] read = new bool[isRead.Length];
            for (int i = 0; i < readFlag.Length; i++)
            {
                // Convert the value to Boolean
                isRead[i] = Convert.ToBoolean(readFlag[i]);

                // Get the opposite value
                read[i] = !isRead[i];
            }

            return read;
        }

        /// <summary>
        /// Updates the subject of the specified item.
        /// </summary>
        /// <param name="itemIds">ID of the item which should be updated.</param>
        /// <param name="itemSubject">The new subject of the specified item.</param>
        protected void UpdateItemSubject(BaseItemIdType[] itemIds, string itemSubject)
        {
            ItemChangeType[] itemChanges = new ItemChangeType[itemIds.Length];

            // Set the public properties (Subject) which all the seven kinds of operation have.
            for (int i = 0; i < itemIds.Length; i++)
            {
                itemChanges[i] = new ItemChangeType();
                itemChanges[i].Item = itemIds[i];
                itemChanges[i].Updates = new ItemChangeDescriptionType[]
                    {
                        new SetItemFieldType()
                        {
                            Item = new PathToUnindexedFieldType()
                            {
                                FieldURI = UnindexedFieldURIType.itemSubject
                            },
                            Item1 = new ItemType()
                            {
                               Subject = itemSubject
                            }
                        }
                    };
            }

            // Update the item.
            this.UpdateItem(itemChanges);
        }

        /// <summary>
        /// Deletes the specified item.
        /// </summary>
        /// <param name="itemIds">ID of the item which should be deleted.</param>
        protected void DeleteItem(BaseItemIdType[] itemIds)
        {
            // Create a request for the DeleteItem operation.
            DeleteItemType deleteItemRequest = new DeleteItemType();

            // Delete the master task and all recurring tasks that are associated with the master task.
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;

            // The item is permanently removed from the store.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            // Do not send meeting cancellations.
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            deleteItemRequest.SendMeetingCancellationsSpecified = true;

            // Specify the item to be deleted.
            deleteItemRequest.ItemIds = itemIds;

            // Invoke the delete item operation and get the response.
            DeleteItemResponseType response = this.COREAdapter.DeleteItem(deleteItemRequest);

            // Check whether the DeleteItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Delete item should be successful! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }
        }

        /// <summary>
        /// Updates the specified item.
        /// </summary>
        /// <param name="itemChanges">The array of item changes.</param>
        protected void UpdateItem(ItemChangeType[] itemChanges)
        {
            // Create a request for the UpdateItem operation.
            UpdateItemType request = new UpdateItemType();

            // Add the item and its change to the request.
            request.ItemChanges = itemChanges;

            // The UpdateItem operation automatically resolves any conflict.
            request.ConflictResolution = ConflictResolutionType.AutoResolve;

            // Save the message after updated.
            request.MessageDisposition = MessageDispositionType.SaveOnly;
            request.MessageDispositionSpecified = true;

            // Do not send meeting invitations or cancellations.
            request.SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType.SendToNone;
            request.SendMeetingInvitationsOrCancellationsSpecified = true;

            // Invoke the UpdateItem operation to get the response.
            UpdateItemResponseType response = this.COREAdapter.UpdateItem(request);

            // Check whether the UpdateItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Update item should be successful! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }
        }

        /// <summary>
        /// Creates a meeting message.
        /// </summary>
        /// <param name="createItemRequest">The request of the meeting message to be created.</param>
        /// <returns>The ID of the created meeting item.</returns>
        protected ItemIdType CreateMeetingMessage(CreateItemType createItemRequest)
        {
            // Initialize the instance of ItemIdType.
            ItemIdType itemId = null;

            // Invoke the CreateItem operation to get the response.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

            if (createItemResponse != null && createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                ItemInfoResponseMessageType info = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
                ItemType result = info.Items.Items[0] as ItemType;

                // Get the ID of the created meeting item.
                itemId = result.ItemId;
            }

            return itemId;
        }
        #endregion

        #region Manage multiple folders
        /// <summary>
        /// Log on with specified user account(userName, userPassword, userDomain) and create multiple folders under the specified parent folder
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderNames">An array of parent folders.</param>
        /// <param name="firstLevelFolderName">Name of the first level folder which should be created.</param>
        /// <param name="secondLevelFolderName">Name of the second level folder which should be created.</param>
        /// <param name="searchText">Search text of the search folder which should be created.</param>
        protected void CreateMultipleFolders(
            string userName,
            string userPassword,
            string userDomain,
            List<DistinguishedFolderIdNameType> folderNames,
            string firstLevelFolderName,
            string secondLevelFolderName,
            string searchText)
        {
            bool isSubFolderCreated = false;
            foreach (DistinguishedFolderIdNameType folder in folderNames)
            {
                if (folder == DistinguishedFolderIdNameType.searchfolders)
                {
                    isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSearchFolder(userName, userPassword, userDomain, firstLevelFolderName, searchText);
                }
                else
                {
                    isSubFolderCreated = this.FOLDSUTControlAdapter.CreateSubFolders(userName, userPassword, userDomain, folder.ToString(), firstLevelFolderName, secondLevelFolderName);
                }

                Site.Assert.IsTrue(isSubFolderCreated, string.Format("The new sub folders in '{0}' should be created successfully.", folder));
            }
        }

        /// <summary>
        /// Update multiple folders under the specified parent folder
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderNames">An array of parent folders.</param>
        /// <param name="firstLevelFolderName">Name of the first level folder which should be updated.</param>
        /// <param name="secondLevelFolderName">Name of the second level folder which should be updated.</param>
        /// <param name="newFolderName">A new folder name.</param>
        protected void UpdateMultipleFolders(
            string userName,
            string userPassword,
            string userDomain,
            List<DistinguishedFolderIdNameType> folderNames,
            string firstLevelFolderName,
            string secondLevelFolderName,
            string newFolderName)
        {
            bool updateFolder = false;
            foreach (DistinguishedFolderIdNameType folder in folderNames)
            {
                if (folder == DistinguishedFolderIdNameType.searchfolders)
                {
                    updateFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                        userName,
                        userPassword,
                        userDomain,
                        folder.ToString(),
                        firstLevelFolderName,
                        newFolderName);
                }
                else
                {
                    updateFolder = this.SYNCSUTControlAdapter.FindAndUpdateFolderName(
                        userName,
                        userPassword,
                        userDomain,
                        folder.ToString(),
                        firstLevelFolderName,
                        newFolderName);
                }

                Site.Assert.IsTrue(
                    updateFolder,
                    string.Format("The folder name '{0}' should be updated to '{1}'.", secondLevelFolderName, newFolderName));
            }
        }

        /// <summary>
        /// Log on with specified user account and delete multiple folders under the specified parent folder
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderNames">An array of parent folders.</param>
        /// <param name="newFolderName">Name of the folder which should be deleted.</param>
        protected void DeleteMultipleFolders(string userName, string userPassword, string userDomain, List<DistinguishedFolderIdNameType> folderNames, string newFolderName)
        {
            bool isSubFoldersDeleted = false;
            foreach (DistinguishedFolderIdNameType folder in folderNames)
            {
                isSubFoldersDeleted = this.SYNCSUTControlAdapter.FindAndDeleteSubFolder(
                    userName,
                    userPassword,
                    userDomain,
                    folder.ToString(),
                    newFolderName);

                Site.Assert.IsTrue(
                    isSubFoldersDeleted,
                    string.Format("The folder named '{0}' should be deleted from '{1}' successfully.", newFolderName, folder));
            }
        }
        #endregion

        #region Manage multiple items
        /// <summary>
        /// Creates multiple items
        /// </summary>
        /// <returns>An array of item IDs.</returns>
        protected BaseItemIdType[] CreateMultipleItems()
        {
            // Initialize ItemType list
            List<ItemType> itemType = new List<ItemType>();

            // Get the item type that based on ItemType
            object obj;
            Assembly assembly = Assembly.GetAssembly(typeof(ItemType));
            Type[] types = assembly.GetTypes();

            // Initialize properties that all items have
            PropertyInfo subjectField;
            PropertyInfo fileAs;
            string currentSubject = Common.GenerateResourceName(this.Site, "ItemSubject");
            foreach (Type type in types)
            {
                if (type.BaseType == typeof(ItemType) || (type == typeof(ItemType) && !type.IsAbstract))
                {
                    if (!Common.IsRequirementEnabled(37811, this.Site) && type == typeof(DistributionListType))
                    {
                        // Exchange 2007 does not support DistributionListType item, for detailed information refer to MS-OXWSDLIST.
                        continue;
                    }
                    else if (type == typeof(AbchPersonItemType) && !Common.IsRequirementEnabled(37811006, this.Site))
                    {
                        //Exchange 2007 2010 2013 not support AbchPersonItemType.
                        continue;
                    }
                    else if (type == typeof(RoleMemberItemType) && !Common.IsRequirementEnabled(1752001, this.Site))
                    {
                        //Exchange 2007 2010 2013 not support RoleMemberItemType.
                        continue;
                    }
                    else if (type == typeof(NetworkItemType) && !Common.IsRequirementEnabled(1752003, this.Site))
                    {
                        //Exchange 2007 2010 2013 not support NetworkItemType.
                        continue;
                    }
                    else
                    {
                        string typeName = type.ToString();
                        obj = assembly.CreateInstance(typeName);
                        if (type == typeof(ContactItemType))
                        {
                            fileAs = type.GetProperty("FileAs");
                            fileAs.SetValue(obj, currentSubject, null);
                        }
                        else
                        {
                            subjectField = type.GetProperty("Subject");
                            if (subjectField != null)
                            {
                                subjectField.SetValue(obj, currentSubject, null);
                            }
                        }
                    }

                    itemType.Add((ItemType)obj);
                }
            }

            ItemType[] items = itemType.ToArray();
            CreateItemType createRequest = new CreateItemType()
            {
                Items = new NonEmptyArrayOfAllItemsType() { Items = items },
            };

            createRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            createRequest.MessageDispositionSpecified = true;
            createRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
            createRequest.SendMeetingInvitationsSpecified = true;

            // Call CreateItem to create multiple items on the server.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createRequest);

            // Check whether the CreateItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in createItemResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Create each type of items should not fail! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }

            // Get the item ids in CreateItem response.
            BaseItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            return itemIds;
        }
        #endregion

        #region Common method
        /// <summary>
        /// Verify SyncFoldersItems operation with/without all optional elements.
        /// </summary>
        /// <param name="request">The request of SyncFolderItems operation.</param>
        /// <param name="isIgnoreElementIncluded">A Boolean value indicates whether the request includes Ignore element.</param>
        protected void VerifySyncFolderItemsOperation(SyncFolderItemsType[] request, bool isIgnoreElementIncluded)
        {
            // Invokes SyncFolderItems operation to get initial syncState.
            SyncFolderItemsResponseType[] response = new SyncFolderItemsResponseType[this.FolderIdNameType.Count];
            SyncFolderItemsResponseMessageType[] responseMessages = new SyncFolderItemsResponseMessageType[this.FolderIdNameType.Count];

            // Create multiple items in default folders
            BaseItemIdType[] itemIds = this.CreateMultipleItems();

            // Invokes SyncFolderItems operation to sync the operation result of create items.
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                response[i] = this.SYNCAdapter.SyncFolderItems(request[i]);
                responseMessages[i] = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response[i]);
            }

            // Invokes UpdateItemSubject operation to update the subject of items.
            this.UpdateItemSubject(itemIds, Common.GenerateResourceName(this.Site, "NewItemSubject"));

            // Invokes SyncFolderItems operation to sync the operation result of update item subject.
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                // Assert the changes in response is not null
                Site.Assert.IsNotNull(responseMessages[i].Changes, "There are items created on server, so the changes between server and client should not be null");
                SyncFolderItemsChangesType changes = responseMessages[i].Changes;

                // Assert the Items is not null
                Site.Assert.IsNotNull(changes.Items, "There should be items information returned in SyncFolderItems response since there are items updated on server.");

                // Set the value of Ignore to the first item's Id
                if (isIgnoreElementIncluded)
                {
                    request[i].Ignore = new ItemIdType[1] { (responseMessages[i].Changes.Items[0] as SyncFolderItemsCreateOrUpdateType).Item.ItemId };
                }

                response[i] = this.SYNCAdapter.SyncFolderItems(request[i]);
                responseMessages[i] = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response[i]);
            }

            // Invokes DeleteItem operation to delete items.
            this.DeleteItem(itemIds);

            // Invokes SyncFolderItems operation to sync the operation result of delete item operation.
            for (int i = 0; i < this.FolderIdNameType.Count; i++)
            {
                response[i] = this.SYNCAdapter.SyncFolderItems(request[i]);
                responseMessages[i] = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response[i]);
            }
        }

        /// <summary>
        /// Get the response message.
        /// </summary>
        /// <param name="folder">The folder which should be synchronized.</param>
        /// <param name="responseMessage">A SyncFolderItemsResponseMessageType type message which is gotten from last synchronization.</param>
        /// <param name="defaultShapeNames">Standard sets of properties to return.</param>
        /// <returns>A SyncFolderItemsResponseMessageType type message.</returns>
        protected SyncFolderItemsResponseMessageType GetResponseMessage(
            DistinguishedFolderIdNameType folder,
            SyncFolderItemsResponseMessageType responseMessage,
            DefaultShapeNamesType defaultShapeNames)
        {
            SyncFolderItemsType request = this.CreateSyncFolderItemsRequestWithoutOptionalElements(folder, defaultShapeNames);

            // Assert the SyncState is not null
            Site.Assert.IsNotNull(responseMessage.SyncState, "The synchronization should not be null.");
            request.SyncState = responseMessage.SyncState;

            SyncFolderItemsResponseType response = this.SYNCAdapter.SyncFolderItems(request);
            responseMessage = TestSuiteHelper.EnsureResponse<SyncFolderItemsResponseMessageType>(response);
            return responseMessage;
        }
        #endregion

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        protected void ConfigureSOAPHeader()
        {
            // Set the value of MailboxCulture.
            MailboxCultureType mailboxCulture = new MailboxCultureType();
            string culture = Common.GetConfigurationPropertyValue("MailboxCulture", this.Site);
            mailboxCulture.Value = culture;

            // Set the value of ExchangeImpersonation.
            ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
            impersonation.ConnectingSID = new ConnectingSIDType();
            impersonation.ConnectingSID.Item =
                Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            Dictionary<string, object> headerValues = new Dictionary<string, object>();
            headerValues.Add("MailboxCulture", mailboxCulture);
            headerValues.Add("ExchangeImpersonation", impersonation);
            this.SYNCAdapter.ConfigureSOAPHeader(headerValues);
        }
        #endregion
    }
}