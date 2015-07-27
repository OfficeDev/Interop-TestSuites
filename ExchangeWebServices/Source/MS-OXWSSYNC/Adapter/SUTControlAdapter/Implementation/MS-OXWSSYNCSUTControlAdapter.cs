//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSSYNC SUT control adapter implementation.
    /// </summary>
    public class MS_OXWSSYNCSUTControlAdapter : ManagedAdapterBase, IMS_OXWSSYNCSUTControlAdapter
    {
        #region Fields
        /// <summary>
        /// The endpoint url of Exchange Web Service.
        /// </summary>
        private string url;

        /// <summary>
        /// The password for userName used to access web service.
        /// </summary>
        private string password;

        /// <summary>
        /// The user name used to access web service.
        /// </summary>
        private string userName;

        /// <summary>
        /// The domain of server.
        /// </summary>
        private string domain;

        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSSYNC";

            // Merge configuration files.
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            this.password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSSYNCSUTControlAdapter Operations
        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified meeting message in the Inbox folder, then accept it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="itemSubject">Subject of the meeting message which should be accepted.</param>
        /// <param name="itemType">Type of the item which should be accepted.</param>
        /// <returns>If the specified meeting message is accepted successfully, return true; otherwise, return false.</returns>
        public bool FindAndAcceptMeetingMessage(string userName, string userPassword, string userDomain, string itemSubject, string itemType)
        {
            // Define the Inbox folder as parent folder.
            DistinguishedFolderIdNameType parentFolderIdName = DistinguishedFolderIdNameType.inbox;

            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            Item item = (Item)Enum.Parse(typeof(Item), itemType, true);

            // Loop to find the specified item in the specified folder.
            ItemType type = this.LoopToFindItem(parentFolderIdName, itemSubject, item);
            bool isAccepted = false;
            if (type != null)
            {
                MeetingRequestMessageType message = type as MeetingRequestMessageType;

                // Create a request for the CreateItem operation.
                CreateItemType createItemRequest = new CreateItemType();

                // Add the CalendarItemType item to the items to be created.
                createItemRequest.Items = new NonEmptyArrayOfAllItemsType();

                // Create an AcceptItemType item to reply to a meeting request.
                AcceptItemType acceptItem = new AcceptItemType();

                // Set the related meeting request.
                acceptItem.ReferenceItemId = message.ItemId;
                createItemRequest.Items.Items = new ItemType[] { acceptItem };

                // Set the MessageDisposition property to SendOnly.
                createItemRequest.MessageDisposition = MessageDispositionType.SendOnly;
                createItemRequest.MessageDispositionSpecified = true;

                // Invoke the CreateItem operation.
                CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(createItemRequest);

                if (createItemResponse != null && createItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    isAccepted = true;
                }
            }

            return isAccepted;
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified item then delete it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified item.</param>
        /// <param name="itemSubject">Subject of the item which should be deleted.</param>
        /// <param name="itemType">Type of the item which should be deleted.</param>
        /// <returns>If the specified item is deleted successfully, return true; otherwise, return false.</returns>
        public bool FindAndDeleteItem(string userName, string userPassword, string userDomain, string folderName, string itemSubject, string itemType)
        {
            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Parse the parent folder name to DistinguishedFolderIdNameType.
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            Item item = (Item)Enum.Parse(typeof(Item), itemType, true);
            ItemType type = this.LoopToFindItem(parentFolderIdName, itemSubject, item);

            bool isDeleted = false;
            if (type != null)
            {
                DeleteItemType deleteItemRequest = new DeleteItemType();
                deleteItemRequest.ItemIds = new BaseItemIdType[] { type.ItemId };

                // Invoke the delete item operation and get the response.
                DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);

                if (response != null && response.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    // If delete operation succeeds, return true
                    isDeleted = true;
                }
            }

            return isDeleted;
        }

        /// <summary>
        /// Log on a mailbox with a specified user account and check whether the specified item exists.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified item.</param>
        /// <param name="itemSubject">Subject of the item which should exist.</param>
        /// <param name="itemType">Type of the item which should exist.</param>
        /// <returns>If the item exists, return true; otherwise, return false.</returns>
        public bool IsItemExisting(string userName, string userPassword, string userDomain, string folderName, string itemSubject, string itemType)
        {
            bool isExisting = false;

            // Parse the parent folder name to DistinguishedFolderIdNameType.
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            Item item = (Item)Enum.Parse(typeof(Item), itemType, true);

            // Loop to find the specified item
            ItemType type = this.LoopToFindItem(parentFolderIdName, itemSubject, item);

            if (type != null)
            {
                switch (item)
                {
                    case Item.MeetingRequest:
                        MeetingRequestMessageType requestMessage = type as MeetingRequestMessageType;
                        if (requestMessage != null)
                        {
                            isExisting = true;
                        }

                        break;
                    case Item.MeetingResponse:
                        MeetingResponseMessageType responseMessage = type as MeetingResponseMessageType;
                        if (responseMessage != null)
                        {
                            isExisting = true;
                        }

                        break;
                    case Item.MeetingCancellation:
                        MeetingCancellationMessageType cancellationMessage = type as MeetingCancellationMessageType;
                        if (cancellationMessage != null)
                        {
                            isExisting = true;
                        }

                        break;
                    case Item.CalendarItem:
                        CalendarItemType calendarItem = type as CalendarItemType;
                        if (calendarItem != null)
                        {
                            isExisting = true;
                        }

                        break;
                }
            }

            return isExisting;
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and check whether the specified calendar item is cancelled or not. 
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified meeting message.</param>
        /// <param name="itemSubject">Subject of the meeting message which should exist.</param>
        /// <returns>If the specified calendar item exists and is canceled, return true, otherwise return false.</returns>
        public bool IsCalendarItemCanceled(string userName, string userPassword, string userDomain, string folderName, string itemSubject)
        {
            // Parse the parent folder name to DistinguishedFolderIdNameType.
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Loop to find the meeting message
            ItemType item = this.LoopToFindItem(parentFolderIdName, itemSubject, Item.CalendarItem);
            if (item != null)
            {
                CalendarItemType calendarItem = item as CalendarItemType;

                // If the IsCancelled property of the item is true, return true
                if (calendarItem.IsCancelled == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified folder then update the folder name of it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="currentFolderName">Current name of the folder which will be updated.</param>
        /// <param name="newFolderName">New name of the folder which will be updated to.</param>
        /// <returns>If the name of the folder is updated successfully, return true; otherwise, return false.</returns>
        public bool FindAndUpdateFolderName(string userName, string userPassword, string userDomain, string parentFolderName, string currentFolderName, string newFolderName)
        {
            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Parse the parent folder name to DistinguishedFolderIdNameType.
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), parentFolderName, true);

            // Create UpdateFolder request
            UpdateFolderType updateFolderRequest = new UpdateFolderType();
            updateFolderRequest.FolderChanges = new FolderChangeType[1];
            updateFolderRequest.FolderChanges[0] = new FolderChangeType();
            updateFolderRequest.FolderChanges[0].Item = this.FindSubFolder(parentFolderIdName, currentFolderName);

            // Identify the field to update and the value to set for it.
            SetFolderFieldType displayName = new SetFolderFieldType();
            PathToUnindexedFieldType displayNameProp = new PathToUnindexedFieldType();
            displayNameProp.FieldURI = UnindexedFieldURIType.folderDisplayName;
            FolderType updatedFolder = new FolderType();
            updatedFolder.DisplayName = newFolderName;
            displayName.Item = displayNameProp;
            updatedFolder.DisplayName = newFolderName;
            displayName.Item1 = updatedFolder;

            // Add a single element into the array of changes.
            updateFolderRequest.FolderChanges[0].Updates = new FolderChangeDescriptionType[1];
            updateFolderRequest.FolderChanges[0].Updates[0] = displayName;
            bool isFolderNameUpdated = false;

            // Invoke the UpdateFolder operation and get the response.
            UpdateFolderResponseType updateFolderResponse = this.exchangeServiceBinding.UpdateFolder(updateFolderRequest);

            if (updateFolderResponse != null && ResponseClassType.Success == updateFolderResponse.ResponseMessages.Items[0].ResponseClass)
            {
                isFolderNameUpdated = true;
            }

            return isFolderNameUpdated;
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified folder, then delete it if it is found.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="subFolderName">Name of the folder which will be updated.</param>
        /// <returns>If the folder is deleted successfully, return true; otherwise, return false.</returns>
        public bool FindAndDeleteSubFolder(string userName, string userPassword, string userDomain, string parentFolderName, string subFolderName)
        {
            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Parse the parent folder name to DistinguishedFolderIdNameType.
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), parentFolderName, true);

            DeleteFolderType deleteFolderRequest = new DeleteFolderType();
            deleteFolderRequest.DeleteType = DisposalType.HardDelete;
            deleteFolderRequest.FolderIds = new BaseFolderIdType[1];
            deleteFolderRequest.FolderIds[0] = this.FindSubFolder(parentFolderIdName, subFolderName);

            // Invoke the DeleteFolder operation and get the response.
            DeleteFolderResponseType deleteFolderResponse = this.exchangeServiceBinding.DeleteFolder(deleteFolderRequest);

            bool isDeleted = false;
            if (deleteFolderResponse != null && deleteFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                // If the DeleteFolder operation succeeds, return true.
                isDeleted = true;
            }

            return isDeleted;
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and delete all the items and subfolders from Inbox, Sent Items, Calendar, Contacts, Tasks, Deleted Items and Search Folders if any.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <returns>If the mailbox is cleaned up successfully, return true; otherwise, return false.</returns>
        public bool CleanupMailBox(string userName, string userPassword, string userDomain)
        {
            // Switch to specified user mailbox.
            bool isSwitched = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isSwitched,
                string.Format("Logon with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            bool isCleaned = false;
            if (
                   this.CleanupFolder(DistinguishedFolderIdNameType.inbox) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.sentitems) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.calendar) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.contacts) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.tasks) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.deleteditems) &&
                   this.CleanupFolder(DistinguishedFolderIdNameType.searchfolders))
            {
                isCleaned = true;
            }

            return isCleaned;
        }
                   
        #endregion

        #region Private Methods
        /// <summary>
        /// Loop to find the item with the specified subject.
        /// </summary>
        /// <param name="folderName">Name of the specified folder.</param>
        /// <param name="itemSubject">Subject of the specified item.</param>
        /// <param name="itemType">Type of the specified item.</param>
        /// <returns>Item with the specified subject.</returns>
        private ItemType LoopToFindItem(DistinguishedFolderIdNameType folderName, string itemSubject, Item itemType)
        {
            ItemType[] items = null;
            ItemType firstFoundItem = null;
            int sleepTimes = 0;

            // Get the query sleep delay and times from ptfconfig file.
            int queryDelay = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int queryTimes = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            // Loop to find the item, in case that the sent item has still not been received.
            do
            {
                Thread.Sleep(queryDelay);
                items = this.FindAllItems(folderName);
                sleepTimes++;
            }
            while (items == null && sleepTimes < queryTimes);

            ItemType type = null;

            switch (itemType)
            {
                case Item.MeetingRequest:
                    type = new MeetingRequestMessageType();
                    break;
                case Item.MeetingResponse:
                    type = new MeetingResponseMessageType();
                    break;
                case Item.MeetingCancellation:
                    type = new MeetingCancellationMessageType();
                    break;
                case Item.CalendarItem:
                    type = new CalendarItemType();
                    break;
            }

            if (items != null)
            {
                // Find the item with the specified subject and store its ID.
                for (int i = 0; i < items.Length; i++)
                {
                    if (items[i].Subject.Contains(itemSubject) && items[i].GetType().ToString() == type.ToString())
                    {
                        firstFoundItem = items[i];
                        break;
                    }
                }
            }

            return firstFoundItem;
        }

        /// <summary>
        /// Find all the items in the specified folder.
        /// </summary>
        /// <param name="folderName">Name of the specified folder.</param>
        /// <returns>An array of found items.</returns>
        private ItemType[] FindAllItems(DistinguishedFolderIdNameType folderName)
        {
            // Create an array of ItemType.
            ItemType[] items = null;

            // Create an instance of FindItemType.
            FindItemType findItemRequest = new FindItemType();
            findItemRequest.ParentFolderIds = new BaseFolderIdType[1];

            DistinguishedFolderIdType parentFolder = new DistinguishedFolderIdType();
            parentFolder.Id = folderName;
            findItemRequest.ParentFolderIds[0] = parentFolder;

            // Get properties that are defined as the default for the items.
            findItemRequest.ItemShape = new ItemResponseShapeType();
            findItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.Default;

            // Invoke the FindItem operation.
            FindItemResponseType findItemResponse = this.exchangeServiceBinding.FindItem(findItemRequest);

            if (findItemResponse != null && findItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                // Get the found items from the response.
                FindItemResponseMessageType findItemMessage = findItemResponse.ResponseMessages.Items[0] as FindItemResponseMessageType;
                ArrayOfRealItemsType findItems = findItemMessage.RootFolder.Item as ArrayOfRealItemsType;
                items = findItems.Items;
            }

            return items;
        }

        /// <summary>
        /// Find the specified sub folder.
        /// </summary>
        /// <param name="parentFolderName">Name of the specified parent folder.</param>
        /// <param name="subFolderName">Name of the specified sub folder.</param>
        /// <returns>Id of the folder.</returns>
        private FolderIdType FindSubFolder(DistinguishedFolderIdNameType parentFolderName, string subFolderName)
        {
            FolderIdType folderId = new FolderIdType();

            // Find all sub folders in the specified parent folder.
            BaseFolderType[] folders = this.FindAllSubFolders(parentFolderName);
            Site.Assert.IsNotNull(folders, "There should be at least one folder in the '{0}' folder.", parentFolderName);

            // Find the item with the specified subject.
            foreach (BaseFolderType currentFolder in folders)
            {
                if (currentFolder.DisplayName == subFolderName)
                {
                    folderId = currentFolder.FolderId;
                    break;
                }
            }

            Site.Assert.IsNotNull(folderId.Id, "There should be at least one folder with the specified subject '{0}'.", subFolderName);
            return folderId;
        }

        /// <summary>
        /// Find all the sub folders in the specified folder.
        /// </summary>
        /// <param name="parentFolderName">Name of the specified parent folder.</param>
        /// <returns>An array of found sub folders.</returns>
        private BaseFolderType[] FindAllSubFolders(DistinguishedFolderIdNameType parentFolderName)
        {
            // Create an array of BaseFolderType.
            BaseFolderType[] folders = null;

            // Create the request and specify the traversal type.
            FindFolderType findFolderRequest = new FindFolderType();
            findFolderRequest.Traversal = FolderQueryTraversalType.Deep;

            // Define the properties to be returned in the response.
            FolderResponseShapeType responseShape = new FolderResponseShapeType();
            responseShape.BaseShape = DefaultShapeNamesType.Default;
            findFolderRequest.FolderShape = responseShape;

            // Identify which folders to search.
            DistinguishedFolderIdType[] folderIDArray = new DistinguishedFolderIdType[1];
            folderIDArray[0] = new DistinguishedFolderIdType();
            folderIDArray[0].Id = parentFolderName;

            // Add the folders to search to the request.
            findFolderRequest.ParentFolderIds = folderIDArray;

            FindFolderResponseType findFolderResponse = this.exchangeServiceBinding.FindFolder(findFolderRequest);
            FindFolderResponseMessageType findFolderResponseMessageType = new FindFolderResponseMessageType();
            if (findFolderResponse != null && findFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                findFolderResponseMessageType = findFolderResponse.ResponseMessages.Items[0] as FindFolderResponseMessageType;
                folders = findFolderResponseMessageType.RootFolder.Folders;
            }

            return folders;
        }

        /// <summary>
        /// Deletes all items and sub folders from the specified folder.
        /// </summary>
        /// <param name="folderName">Name of the specified parent folder.</param>
        /// <param name="needVerify">If need to verify the subfolders,when false,it will cleanup all the folders without verify</param>
        /// <returns>If the specified folder is cleaned up successfully, return true; otherwise, return false.</returns>
        private bool CleanupFolder(DistinguishedFolderIdNameType folderName, bool needVerify = true)
        {
            bool isAllItemsAndFoldersDeleted = false;
            ItemType[] items = this.FindAllItems(folderName);

            // Create a request for the DeleteItem operation.
            DeleteItemType deleteItemRequest = new DeleteItemType();

            // The item is permanently removed from the store.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            // Do not send meeting cancellations.
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            deleteItemRequest.SendMeetingCancellationsSpecified = true;

            if (items != null)
            {
                foreach (ItemType currentItem in items)
                {
                    if (currentItem.GetType() == typeof(TaskType))
                    {
                        deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
                        deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;
                    }

                    deleteItemRequest.ItemIds = new BaseItemIdType[] { currentItem.ItemId };

                    // Invoke the delete item operation.
                    DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);

                    Site.Assert.AreEqual<ResponseClassType>(
                            ResponseClassType.Success,
                            response.ResponseMessages.Items[0].ResponseClass,
                            "The delete item operation should execute successfully.");
                }
            }

            // Find all sub folders in the specified folder.
            BaseFolderType[] folders = this.FindAllSubFolders(folderName);

            if (folders.Length != 0)
            {
                foreach (BaseFolderType currentFolder in folders)
                {
                    if (needVerify)
                    {
                        bool isCreatedByCase = false;
                        AdapterHelper.CreatedFolders.ForEach(r =>
                        {
                            if (r.FolderId.Id == currentFolder.FolderId.Id)
                            {
                                isCreatedByCase = true;
                            }
                        });

                        if (!isCreatedByCase)
                        {
                            continue;
                        }
                    }

                    FolderIdType responseFolderId = currentFolder.FolderId;

                    FolderIdType folderId = new FolderIdType();
                    folderId.Id = responseFolderId.Id;

                    DeleteFolderType deleteFolderRequest = new DeleteFolderType();
                    deleteFolderRequest.DeleteType = DisposalType.HardDelete;
                    deleteFolderRequest.FolderIds = new BaseFolderIdType[1];
                    deleteFolderRequest.FolderIds[0] = folderId;

                    // Send the request and get the response.
                    DeleteFolderResponseType deleteFolderResponse = this.exchangeServiceBinding.DeleteFolder(deleteFolderRequest);

                    // Delete folder operation should return response info.
                    if (deleteFolderResponse.ResponseMessages.Items[0] != null)
                    {
                        Site.Assert.AreEqual<ResponseClassType>(
                            ResponseClassType.Success,
                            deleteFolderResponse.ResponseMessages.Items[0].ResponseClass,
                            "The delete folder operation should be successful.");
                    }
                }
            }

            // Invoke the FindItem operation.
            items = this.FindAllItems(folderName);

            // Invoke the FindFolder operation.
            folders = this.FindAllSubFolders(folderName);

            // If neither items and sub folders could be found, the folder has been cleaned up successfully.
            bool allFoldersBelongstoSys = true;

            foreach (BaseFolderType folder in folders)
            {               
                bool find = false;
                AdapterHelper.CreatedFolders.ForEach(r =>
                {
                    if (r.FolderId.Id == folder.FolderId.Id)
                    {
                        find = true;
                    }
                });

                if (find)
                {
                    allFoldersBelongstoSys = false;
                    break;
                }
            }

            if (items == null && (folders.Length == 0 || allFoldersBelongstoSys))
            {
                isAllItemsAndFoldersDeleted = true;
            }

            return isAllItemsAndFoldersDeleted;
        }

        #endregion
    }
}