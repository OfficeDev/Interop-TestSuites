//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Fields

        /// <summary>
        /// Specify if a permission authorized user can create subfolder.
        /// </summary>
        private bool canCreateSubFolder;

        /// <summary>
        /// Specify if a permission authorized user can read subfolder.
        /// </summary>
        private bool canReadSubFolder;

        /// <summary>
        /// Specify if a permission authorized user can edit subfolder.
        /// </summary>
        private bool canEditSubFolder;

        /// <summary>
        /// Specify if a permission authorized user can delete subfolder.
        /// </summary>
        private bool canDeleteSubFolder;

        /// <summary>
        /// Specify if a permission authorized user can create item.
        /// </summary>
        private bool canCreateItem;

        /// <summary>
        ///  Specify if a permission authorized user can read items which is created by the user.
        /// </summary>
        private bool canReadOwnedItem;

        /// <summary>
        /// Specify if a permission authorized user can edit items which is created by the user.
        /// </summary>
        private bool canEditOwnedItem;

        /// <summary>
        /// Specify if a permission authorized user can delete items which is created by the user.
        /// </summary>
        private bool canDeleteOwnedItem;

        /// <summary>
        ///  Specify if a permission authorized user can read items which isn't created by the user.
        /// </summary>
        private bool canReadNotOwnedItem;

        /// <summary>
        ///  Specify if a permission authorized user can edit items which isn't created by the user.
        /// </summary>
        private bool canEditNotOwnedItem;

        /// <summary>
        /// Specify if a permission authorized user can delete items which isn't created by the user.
        /// </summary>
        private bool canDeleteNotOwnedItem;

        /// <summary>
        /// Variable to save the new created folder's folder Ids.
        /// </summary>
        private Collection<FolderIdType> newCreatedFolderIds;

        /// <summary>
        /// Variable to save the new created items' Ids.
        /// </summary>
        private Collection<ItemIdType> newCreatedItemIds;

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets a value indicating whether a subfolder can be created..
        /// </summary>
        protected bool CanCreateSubFolder
        {
            get { return this.canCreateSubFolder; }
            set { this.canCreateSubFolder = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a subfolder can be read.
        /// </summary>
        protected bool CanReadSubFolder
        {
            get { return this.canReadSubFolder; }
            set { this.canReadSubFolder = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a subfolder can be edited.
        /// </summary>
        protected bool CanEditSubFolder
        {
            get { return this.canEditSubFolder; }
            set { this.canEditSubFolder = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a subfolder can be deleted.
        /// </summary>
        protected bool CanDeleteSubFolder
        {
            get { return this.canDeleteSubFolder; }
            set { this.canDeleteSubFolder = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item can be created.
        /// </summary>
        protected bool CanCreateItem
        {
            get { return this.canCreateItem; }
            set { this.canCreateItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item owned by a user can be read.
        /// </summary>
        protected bool CanReadOwnedItem
        {
            get { return this.canReadOwnedItem; }
            set { this.canReadOwnedItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item owned by a user can be edited.
        /// </summary>
        protected bool CanEditOwnedItem
        {
            get { return this.canEditOwnedItem; }
            set { this.canEditOwnedItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item owned by a user can be deleted.
        /// </summary>
        protected bool CanDeleteOwnedItem
        {
            get { return this.canDeleteOwnedItem; }
            set { this.canDeleteOwnedItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item not owned by a user can be read.
        /// </summary>
        protected bool CanReadNotOwnedItem
        {
            get { return this.canReadNotOwnedItem; }
            set { this.canReadNotOwnedItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item not owned by a user can be edited.
        /// </summary>
        protected bool CanEditNotOwnedItem
        {
            get { return this.canEditNotOwnedItem; }
            set { this.canEditNotOwnedItem = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether an item not owned by a user can be deleted.
        /// </summary>
        protected bool CanDeleteNotOwnedItem
        {
            get { return this.canDeleteNotOwnedItem; }
            set { this.canDeleteNotOwnedItem = value; }
        }

        /// <summary>
        /// Gets newCreatedFolderIds which contains created folder ids.
        /// </summary>
        protected Collection<FolderIdType> NewCreatedFolderIds
        {
            get { return this.newCreatedFolderIds; }
        }

        /// <summary>
        /// Gets newCreatedItemIds which contains created folder ids.
        /// </summary>
        protected Collection<ItemIdType> NewCreatedItemIds
        {
            get { return this.newCreatedItemIds; }
        }

        /// <summary>
        /// Gets MS-OXWSFOLD protocol adapter.
        /// </summary>
        protected IMS_OXWSFOLDAdapter FOLDAdapter { get; private set; }

        /// <summary>
        /// Gets MS-OXWSCORE protocol adapter.
        /// </summary>
        protected IMS_OXWSCOREAdapter COREAdapter { get; private set; }

        /// <summary>
        /// Gets MS-OXWSSRCH protocol adapter.
        /// </summary>
        protected IMS_OXWSSRCHAdapter SRCHAdapter { get; private set; }

        #endregion

        #region Test case initialize and clean up

        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            // Following codes shall be run before every test case execution.
            base.TestInitialize();
            this.newCreatedFolderIds = new Collection<FolderIdType>();
            this.newCreatedItemIds = new Collection<ItemIdType>();
            this.FOLDAdapter = Site.GetAdapter<IMS_OXWSFOLDAdapter>();
            this.COREAdapter = Site.GetAdapter<IMS_OXWSCOREAdapter>();
            this.SRCHAdapter = Site.GetAdapter<IMS_OXWSSRCHAdapter>();
            this.InitialPermissionVariables();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // Ensure use the right user to clean up.
            #region Switch user

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            // Following codes shall be run after every test case execution.
            #region Delete the new created folder to make sure each case run at the same environment.

            // Clean up all created folders.
            if (this.newCreatedFolderIds.Count > 0)
            {
                for (int temp = 0; temp < this.newCreatedFolderIds.Count; temp++)
                {
                    // Delete folder request.
                    DeleteFolderType deleteFolderRequest = new DeleteFolderType();

                    // Specify the delete type.
                    deleteFolderRequest.DeleteType = DisposalType.HardDelete;

                    // Set the deleteFolderRequest's folderId type.
                    deleteFolderRequest.FolderIds = new BaseFolderIdType[1];
                    deleteFolderRequest.FolderIds[0] = this.newCreatedFolderIds[temp];

                    // Delete the specified folder.
                    this.FOLDAdapter.DeleteFolder(deleteFolderRequest);
                }
            }

            // Clean up all created items.
            if (this.newCreatedItemIds.Count > 0)
            {
                for (int temp = 0; temp < this.newCreatedItemIds.Count; temp++)
                {
                    this.DeleteItem(this.newCreatedItemIds[temp]);
                }
            }

            #endregion
        }

        #endregion

        #region Test case base methods

        /// <summary>
        /// Create item within a specific folder.
        /// </summary>
        /// <param name="toAddress">To address of created item</param>
        /// <param name="folderId">Parent folder id of the created item.</param>
        /// <param name="subject">Subject of the item.</param>
        /// <returns>Id of created item.</returns>
        protected ItemIdType CreateItem(string toAddress, string folderId, string subject)
        {
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.MessageDispositionSpecified = true;
            createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
            createItemRequest.SendMeetingInvitationsSpecified = true;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();

            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType distinguishedFolderIdName = new DistinguishedFolderIdNameType();
            bool isSuccess = Enum.TryParse<DistinguishedFolderIdNameType>(folderId, true, out distinguishedFolderIdName);

            if (isSuccess)
            {
                distinguishedFolderId.Id = distinguishedFolderIdName;
                createItemRequest.SavedItemFolderId.Item = distinguishedFolderId;
            }
            else
            {
                FolderIdType id = new FolderIdType();
                id.Id = folderId;
                createItemRequest.SavedItemFolderId.Item = id;
            }

            MessageType message = new MessageType();
            message.Subject = subject;
            EmailAddressType address = new EmailAddressType();
            address.EmailAddress = toAddress;

            // Set this message to unread.
            message.IsRead = false;
            message.IsReadSpecified = true;
            message.ToRecipients = new EmailAddressType[1];
            message.ToRecipients[0] = address;
            BodyType body = new BodyType();
            body.Value = Common.GenerateResourceName(this.Site, "Test Mail Body");
            message.Body = body;

            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[1];
            createItemRequest.Items.Items[0] = message;
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            ItemInfoResponseMessageType itemInfo = (ItemInfoResponseMessageType)createItemResponse.ResponseMessages.Items[0];

            // Return item id.
            if (itemInfo.Items.Items != null)
            {
                return itemInfo.Items.Items[0].ItemId;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Find item within a specific folder.
        /// </summary>
        /// <param name="folderName">The name of the folder to search item.</param>
        /// <param name="itemSubject">The subject of the item to be searched.</param>
        /// <returns>Id of found item.</returns>
        protected ItemIdType FindItem(string folderName, string itemSubject)
        {
            // Create the request and specify the parent folder ID.
            FindItemType findItemRequest = new FindItemType();
            findItemRequest.ParentFolderIds = new BaseFolderIdType[1];

            DistinguishedFolderIdType parentFolder = new DistinguishedFolderIdType();
            parentFolder.Id = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            findItemRequest.ParentFolderIds[0] = parentFolder;

            // Get properties that are defined as the default for the items.
            findItemRequest.ItemShape = new ItemResponseShapeType();
            findItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // Await the item created properly.
            int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int timesOfSleep = 0;
            ItemIdType itemId = null;
            do
            {
                Thread.Sleep(sleepTime);

                // Invoke the FindItem operation.
                FindItemResponseType findItemResponse = this.SRCHAdapter.FindItem(findItemRequest);

                if (findItemResponse != null)
                {
                    // Get the found items from the response.
                    FindItemResponseMessageType findItemMessage = findItemResponse.ResponseMessages.Items[0] as FindItemResponseMessageType;
                    ArrayOfRealItemsType itemArray = findItemMessage.RootFolder.Item as ArrayOfRealItemsType;
                    ItemType[] items = itemArray.Items;
                    if (items != null)
                    {
                        foreach (ItemType item in items)
                        {
                            if (item.Subject == itemSubject)
                            {
                                itemId = item.ItemId;
                                return itemId;
                            }
                        }
                    }

                    timesOfSleep++;
                }
            }
            while (itemId == null && timesOfSleep <= retryCount);
            return null;
        }

        /// <summary>
        /// Find if item within a specific folder is deleted.
        /// </summary>
        /// <param name="folderName">The name of the folder to search item.</param>
        /// <param name="itemSubject">The subject of the item to be searched.</param>
        /// <returns>If item has been deleted.</returns>
        protected bool IfItemDeleted(string folderName, string itemSubject)
        {
            // Create the request and specify the parent folder ID.
            FindItemType findItemRequest = new FindItemType();
            findItemRequest.ParentFolderIds = new BaseFolderIdType[1];

            DistinguishedFolderIdType parentFolder = new DistinguishedFolderIdType();
            parentFolder.Id = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), folderName, true);

            findItemRequest.ParentFolderIds[0] = parentFolder;

            // Get properties that are defined as the default for the items.
            findItemRequest.ItemShape = new ItemResponseShapeType();
            findItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // Await the item created properly.
            int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int timesOfSleep = 0;
            bool itemDeleted = true;
            do
            {
                itemDeleted = true;
                Thread.Sleep(sleepTime);

                // Invoke the FindItem operation.
                FindItemResponseType findItemResponse = this.SRCHAdapter.FindItem(findItemRequest);

                if (findItemResponse != null)
                {
                    // Get the found items from the response.
                    FindItemResponseMessageType findItemMessage = findItemResponse.ResponseMessages.Items[0] as FindItemResponseMessageType;
                    ArrayOfRealItemsType itemArray = findItemMessage.RootFolder.Item as ArrayOfRealItemsType;
                    ItemType[] items = itemArray.Items;
                    if (items != null)
                    {
                        foreach (ItemType item in items)
                        {
                            if (item.Subject == itemSubject)
                            {
                                itemDeleted = false;
                                break;
                            }
                        }
                    }

                    timesOfSleep++;
                }
            }
            while (!itemDeleted && timesOfSleep <= retryCount);

            return itemDeleted;
        }

        /// <summary>
        /// Update subject of specific items.
        /// </summary>
        /// <param name="itemIds">An array of folder identifiers.</param>
        /// <returns>If item updated successfully.</returns>
        protected bool UpdateItemSubject(params ItemIdType[] itemIds)
        {
            UpdateItemType updateRequest = new UpdateItemType();
            ItemChangeType[] itemChanges = new ItemChangeType[itemIds.Length];

            for (int index = 0; index < itemIds.Length; index++)
            {
                itemChanges[index] = new ItemChangeType();
                itemChanges[index].Item = itemIds[index];
                itemChanges[index].Updates = new ItemChangeDescriptionType[1];
                SetItemFieldType setItem = new SetItemFieldType();
                setItem.Item = new PathToUnindexedFieldType()
                {
                    FieldURI = UnindexedFieldURIType.itemSubject
                };
                setItem.Item1 = new ContactItemType()
                {
                    Subject = Common.GenerateResourceName(this.Site, "ItemSubjectUpdated", (uint)index),
                };
                itemChanges[index].Updates[0] = setItem;
            }

            updateRequest.ItemChanges = itemChanges;
            updateRequest.MessageDispositionSpecified = true;
            updateRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            updateRequest.SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType.SendToAllAndSaveCopy;
            updateRequest.SendMeetingInvitationsOrCancellationsSpecified = true;

            UpdateItemResponseType updateItemResponse = this.COREAdapter.UpdateItem(updateRequest);

            // A Boolean indicates whether the response is a success.
            bool isSuccess = new bool();

            for (int index = 0; index < itemIds.Length; index++)
            {
                isSuccess = ResponseClassType.Success == updateItemResponse.ResponseMessages.Items[index].ResponseClass;

                if (isSuccess)
                {
                    continue;
                }
                else
                {
                    break;
                }
            }

            return isSuccess;
        }

        /// <summary>
        /// Get information of specific items.
        /// </summary>
        /// <param name="itemIds">An array of folder identifiers.</param>
        /// <returns>If item information returned successfully.</returns>
        protected bool GetItem(params ItemIdType[] itemIds)
        {
            GetItemType getItemRequest = new GetItemType();

            // The Items properties returned
            getItemRequest.ItemShape = new ItemResponseShapeType();
            getItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // The items to get
            getItemRequest.ItemIds = itemIds;
            GetItemResponseType getItemResponse = this.COREAdapter.GetItem(getItemRequest);
            return ResponseClassType.Success == getItemResponse.ResponseMessages.Items[0].ResponseClass;
        }

        /// <summary>
        /// Delete specific item.
        /// </summary>
        /// <param name="itemId">Id of specific item.</param>
        /// <returns>If specific item deleted successfully.</returns>
        protected bool DeleteItem(ItemIdType itemId)
        {
            // If id is null return false.
            if (itemId == null)
            {
                return false;
            }

            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            deleteItemRequest.SendMeetingCancellationsSpecified = true;

            // Serialize item ids and change keys to ItemIdType arrays.
            ItemIdType[] itemIdTypes = new ItemIdType[1];
            itemIdTypes[0] = itemId;

            deleteItemRequest.ItemIds = itemIdTypes;
            deleteItemRequest.DeleteType = DisposalType.HardDelete;
            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);
            return deleteItemResponse.ResponseMessages.Items[0].ResponseCode == ResponseCodeType.NoError;
        }

        /// <summary>
        /// Set related folder properties of create folder request
        /// </summary>
        /// <param name="displayNames">Display names of folders that will be set into create folder request.</param>
        /// <param name="folderClasses">Folder class values of folders that will be set into create folder request.</param>
        /// <param name="folderPermissions">Folder permission values of folders that will be set into create folder request. </param>
        /// <param name="createFolderRequest">Create folder request instance that needs to set property values.</param>
        /// <returns>Create folder request instance that have folder property value configured.</returns>
        protected CreateFolderType ConfigureFolderProperty(string[] displayNames, string[] folderClasses, PermissionSetType[] folderPermissions, CreateFolderType createFolderRequest)
        {
            Site.Assert.IsNotNull(displayNames, "Display names should not be null!");
            Site.Assert.IsNotNull(folderClasses, "Folder classes should not be null!");
            Site.Assert.AreEqual<int>(displayNames.Length, folderClasses.Length, "Folder names count should equals to folder class value count!");
            if (folderPermissions != null)
            {
                Site.Assert.AreEqual<int>(displayNames.Length, folderPermissions.Length, "Folder names count should equals to folder permission value count!");
            }

            int folderCount = displayNames.Length;
            createFolderRequest.Folders = new BaseFolderType[folderCount];
            for (int folderPropertyIndex = 0; folderPropertyIndex < folderCount; folderPropertyIndex++)
            {
                string folderResourceName = Common.GenerateResourceName(this.Site, displayNames[folderPropertyIndex]);

                if (folderClasses[folderPropertyIndex] == "IPF.Appointment")
                {
                    CalendarFolderType calendarFolder = new CalendarFolderType();
                    calendarFolder.DisplayName = folderResourceName;
                    createFolderRequest.Folders[folderPropertyIndex] = calendarFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Contact")
                {
                    ContactsFolderType contactFolder = new ContactsFolderType();
                    contactFolder.DisplayName = folderResourceName;
                    if (folderPermissions != null)
                    {
                        contactFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = contactFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Task")
                {
                    TasksFolderType taskFolder = new TasksFolderType();
                    taskFolder.DisplayName = folderResourceName;
                    if (folderPermissions != null)
                    {
                        taskFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = taskFolder;
                }
                else if (folderClasses[folderPropertyIndex] == "IPF.Search")
                {
                    SearchFolderType searchFolder = new SearchFolderType();
                    searchFolder.DisplayName = folderResourceName;

                    // Set search parameters.
                    searchFolder.SearchParameters = new SearchParametersType();
                    searchFolder.SearchParameters.Traversal = SearchFolderTraversalType.Deep;
                    searchFolder.SearchParameters.TraversalSpecified = true;
                    searchFolder.SearchParameters.BaseFolderIds = new DistinguishedFolderIdType[1];
                    DistinguishedFolderIdType inboxType = new DistinguishedFolderIdType();
                    inboxType.Id = new DistinguishedFolderIdNameType();
                    inboxType.Id = DistinguishedFolderIdNameType.inbox;
                    searchFolder.SearchParameters.BaseFolderIds[0] = inboxType;

                    // Use the following search filter 
                    searchFolder.SearchParameters.Restriction = new RestrictionType();
                    PathToUnindexedFieldType path = new PathToUnindexedFieldType();
                    path.FieldURI = UnindexedFieldURIType.itemSubject;
                    RestrictionType restriction = new RestrictionType();
                    ExistsType isEqual = new ExistsType();
                    isEqual.Item = path;
                    restriction.Item = isEqual;
                    searchFolder.SearchParameters.Restriction = restriction;

                    if (folderPermissions != null)
                    {
                        searchFolder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = searchFolder;
                }
                else
                {
                    // Set Display Name and Folder Class for the folder to be created.
                    FolderType folder = new FolderType();
                    folder.DisplayName = folderResourceName;
                    folder.FolderClass = folderClasses[folderPropertyIndex];

                    if (folderPermissions != null)
                    {
                        folder.PermissionSet = folderPermissions[folderPropertyIndex];
                    }

                    createFolderRequest.Folders[folderPropertyIndex] = folder;
                }
            }

            return createFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "CreateFolder".
        /// </summary>
        /// <param name="parentFolderId">The folder identifier for the parent folder.</param>
        /// <param name="folderNames">An array of display name of the folders to be created.</param>
        /// <param name="folderClasses">An array of folder class value of the folders to be created.</param>
        /// <param name="permissionSet">An array of permission set value of the folder.</param>
        /// <returns>Create folder request instance that will send to server.</returns>
        protected CreateFolderType GetCreateFolderRequest(string parentFolderId, string[] folderNames, string[] folderClasses, PermissionSetType[] permissionSet)
        {
            CreateFolderType createFolderRequest = new CreateFolderType();
            createFolderRequest.ParentFolderId = new TargetFolderIdType();

            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType distinguishedFolderIdName = new DistinguishedFolderIdNameType();
            bool isSuccess = Enum.TryParse<DistinguishedFolderIdNameType>(parentFolderId, true, out distinguishedFolderIdName);

            if (isSuccess)
            {
                distinguishedFolderId.Id = distinguishedFolderIdName;
                createFolderRequest.ParentFolderId.Item = distinguishedFolderId;
            }
            else
            {
                FolderIdType id = new FolderIdType();
                id.Id = parentFolderId;
                createFolderRequest.ParentFolderId.Item = id;
            }

            createFolderRequest = this.ConfigureFolderProperty(folderNames, folderClasses, permissionSet, createFolderRequest);

            return createFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "GetFolder".
        /// </summary>
        /// <param name="shapeName">The properties to include in the response.</param>
        /// <param name="folderIds">An array of folder identifiers.</param>   
        /// <returns>Get folder request instance that will send to server.</returns>
        protected GetFolderType GetGetFolderRequest(DefaultShapeNamesType shapeName, params BaseFolderIdType[] folderIds)
        {
            Site.Assert.IsNotNull(folderIds, "Folders id should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderIds.Length, "Folders id should contains at least one Id!");
            GetFolderType getFolderRequest = new GetFolderType();

            // Specify how many folders need to be gotten.
            int folderCount = folderIds.Length;

            // Set the request's folderId.
            getFolderRequest.FolderIds = new BaseFolderIdType[folderCount];

            for (int folderIdIndex = 0; folderIdIndex < folderCount; folderIdIndex++)
            {
                getFolderRequest.FolderIds[folderIdIndex] = folderIds[folderIdIndex];
            }

            // Set folder shape.
            getFolderRequest.FolderShape = new FolderResponseShapeType();
            getFolderRequest.FolderShape.BaseShape = shapeName;
            return getFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "DeleteFolder".
        /// </summary>
        /// <param name="deleteType">How folders are to be deleted.</param>
        /// <param name="folderIds">An array of folder identifier of the folders need to be deleted</param>
        /// <returns>Delete folder request instance that will send to server.</returns>
        protected DeleteFolderType GetDeleteFolderRequest(DisposalType deleteType, params BaseFolderIdType[] folderIds)
        {
            Site.Assert.IsNotNull(folderIds, "Folders id should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderIds.Length, "Folders id should contains at least one Id!");
            DeleteFolderType deleteFolderRequest = new DeleteFolderType();

            // Specify the delete type.
            deleteFolderRequest.DeleteType = deleteType;
            int folderCount = folderIds.Length;

            // Set the request's folderId field.
            deleteFolderRequest.FolderIds = new BaseFolderIdType[folderCount];
            for (int folderIdIndex = 0; folderIdIndex < folderCount; folderIdIndex++)
            {
                deleteFolderRequest.FolderIds[folderIdIndex] = folderIds[folderIdIndex];
            }

            return deleteFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "CreateManagedFolder".
        /// </summary>
        /// <param name="folderNames">An array of names of managed folder.</param>
        /// <returns>Create managed folder request instance that will send to server.</returns>
        protected CreateManagedFolderRequestType GetCreateManagedFolderRequest(params string[] folderNames)
        {
            Site.Assert.IsNotNull(folderNames, "Folder names should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderNames.Length, "Folder names should contains at least one name!");
            CreateManagedFolderRequestType createManagedFolderRequest = new CreateManagedFolderRequestType();
            int folderCount = folderNames.Length;

            // Set the new managed folder's name.
            createManagedFolderRequest.FolderNames = new string[folderCount];
            for (int folderNameIndex = 0; folderNameIndex < folderCount; folderNameIndex++)
            {
                createManagedFolderRequest.FolderNames[folderNameIndex] = folderNames[folderNameIndex];
            }

            return createManagedFolderRequest;
        }

        /// <summary>
        /// Empty a specific folder.
        /// </summary>
        /// <param name="folderId">The folder identifier of the folder need to be emptied.</param>
        /// <param name="deleteType">How an item is deleted.</param>
        /// <param name="deleteSubfolder">Indicates whether the subfolders are also to be deleted. </param>
        /// <returns>Empty folder response instance that will send to server.</returns>
        protected EmptyFolderResponseType CallEmptyFolderOperation(BaseFolderIdType folderId, DisposalType deleteType, bool deleteSubfolder)
        {
            // EmptyFolder request.
            EmptyFolderType emptyFolderRequest = new EmptyFolderType();

            // Specify the delete type.
            emptyFolderRequest.DeleteType = deleteType;

            // Specify which folder will be emptied.
            emptyFolderRequest.FolderIds = new BaseFolderIdType[1];
            emptyFolderRequest.FolderIds[0] = folderId;
            emptyFolderRequest.DeleteSubFolders = deleteSubfolder;

            // Empty the specific folder
            EmptyFolderResponseType emptyFolderResponse = this.FOLDAdapter.EmptyFolder(emptyFolderRequest);

            return emptyFolderResponse;
        }

        /// <summary>
        /// Generate the request message for operation "CopyFolder".
        /// </summary>
        /// <param name="toFolderId">A target folder for operations that copy folders.</param>
        /// <param name="folderIds">An array of folder identifier of the folders need to be copied.</param>
        /// <returns>Copy folder request instance that will send to server.</returns>
        protected CopyFolderType GetCopyFolderRequest(string toFolderId, params BaseFolderIdType[] folderIds)
        {
            Site.Assert.IsNotNull(folderIds, "Folders id should not be null!");
            Site.Assert.AreNotEqual<int>(0, folderIds.Length, "Folders id should contains at least one id!");

            // CopyFolder request.
            CopyFolderType copyFolderRequest = new CopyFolderType();
            int folderCount = folderIds.Length;

            // Identify the folders to be copied.
            copyFolderRequest.FolderIds = new BaseFolderIdType[folderCount];
            for (int folderIdIndex = 0; folderIdIndex < folderCount; folderIdIndex++)
            {
                copyFolderRequest.FolderIds[folderIdIndex] = folderIds[folderIdIndex];
            }

            // Identify the destination folder.
            copyFolderRequest.ToFolderId = new TargetFolderIdType();

            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType distinguishedFolderIdName = new DistinguishedFolderIdNameType();
            bool isSuccess = Enum.TryParse<DistinguishedFolderIdNameType>(toFolderId, true, out distinguishedFolderIdName);

            if (isSuccess)
            {
                distinguishedFolderId.Id = distinguishedFolderIdName;
                copyFolderRequest.ToFolderId.Item = distinguishedFolderId;
            }
            else
            {
                FolderIdType id = new FolderIdType();
                id.Id = toFolderId;
                copyFolderRequest.ToFolderId.Item = id;
            }

            return copyFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "UpdateFolder".
        /// </summary>
        /// <param name="folderType">An array of folder types.</param>
        /// <param name="updateType">An array of update folder types.</param>
        /// <param name="folderIds">An array of folder Ids.</param>
        /// <returns>Update folder request instance that will send to server.</returns>
        protected UpdateFolderType GetUpdateFolderRequest(string[] folderType, string[] updateType, FolderIdType[] folderIds)
        {
            Site.Assert.AreEqual<int>(folderType.Length, folderIds.Length, "Folder type count should equal to folder id count!");
            Site.Assert.AreEqual<int>(folderType.Length, updateType.Length, "Folder type count should equal to update type count!");

            // UpdateFolder request.
            UpdateFolderType updateFolderRequest = new UpdateFolderType();
            int folderCount = folderIds.Length;

            // Set the request's folder id field to Custom Folder's folder id.
            updateFolderRequest.FolderChanges = new FolderChangeType[folderCount];
            for (int folderIndex = 0; folderIndex < folderCount; folderIndex++)
            {
                updateFolderRequest.FolderChanges[folderIndex] = new FolderChangeType();
                updateFolderRequest.FolderChanges[folderIndex].Item = folderIds[folderIndex];

                // Add the array of changes; in this case, a single element array.
                updateFolderRequest.FolderChanges[folderIndex].Updates = new FolderChangeDescriptionType[1];

                switch (updateType[folderIndex])
                {
                    case "SetFolderField":
                        {
                            // Set the new folder name of the specific folder.
                            SetFolderFieldType setFolderField = new SetFolderFieldType();
                            PathToUnindexedFieldType displayNameProp = new PathToUnindexedFieldType();
                            displayNameProp.FieldURI = UnindexedFieldURIType.folderDisplayName;

                            switch (folderType[folderIndex])
                            {
                                case "Folder":
                                    FolderType updatedFolder = new FolderType();
                                    updatedFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = updatedFolder;
                                    break;
                                case "CalendarFolder":
                                    CalendarFolderType updatedCalendarFolder = new CalendarFolderType();
                                    updatedCalendarFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = updatedCalendarFolder;
                                    break;
                                case "ContactsFolder":
                                    CalendarFolderType updatedContactFolder = new CalendarFolderType();
                                    updatedContactFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = updatedContactFolder;
                                    break;
                                case "SearchFolder":
                                    CalendarFolderType updatedSearchFolder = new CalendarFolderType();
                                    updatedSearchFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = updatedSearchFolder;
                                    break;
                                case "TasksFolder":
                                    CalendarFolderType updatedTaskFolder = new CalendarFolderType();
                                    updatedTaskFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = updatedTaskFolder;
                                    break;

                                default:
                                    FolderType generalFolder = new FolderType();
                                    generalFolder.DisplayName = Common.GenerateResourceName(this.Site, "UpdatedFolder" + folderIndex);
                                    setFolderField.Item1 = generalFolder;
                                    break;
                            }

                            setFolderField.Item = displayNameProp;
                            updateFolderRequest.FolderChanges[folderIndex].Updates[0] = setFolderField;
                        }

                        break;
                    case "DeleteFolderField":
                        {
                            // Use DeleteFolderFieldType.
                            DeleteFolderFieldType delFolder = new DeleteFolderFieldType();
                            PathToUnindexedFieldType delProp = new PathToUnindexedFieldType();
                            delProp.FieldURI = UnindexedFieldURIType.folderPermissionSet;
                            delFolder.Item = delProp;
                            updateFolderRequest.FolderChanges[folderIndex].Updates[0] = delFolder;
                        }

                        break;
                    case "AppendToFolderField":
                        {
                            // Use AppendToFolderFieldType.
                            AppendToFolderFieldType appendToFolderField = new AppendToFolderFieldType();
                            PathToUnindexedFieldType displayNameAppendTo = new PathToUnindexedFieldType();
                            displayNameAppendTo.FieldURI = UnindexedFieldURIType.calendarAdjacentMeetings;
                            appendToFolderField.Item = displayNameAppendTo;
                            FolderType folderAppendTo = new FolderType();
                            folderAppendTo.FolderId = folderIds[folderIndex];
                            appendToFolderField.Item1 = folderAppendTo;
                            updateFolderRequest.FolderChanges[folderIndex].Updates[0] = appendToFolderField;
                        }

                        break;
                }
            }

            return updateFolderRequest;
        }

        /// <summary>
        /// Generate the request message for operation "UpdateFolder".
        /// </summary>
        /// <param name="folderType">Identifies type of the folder.</param>
        /// <param name="updateType">Identifies update type of the folder.</param>
        /// <param name="folderId">The folder identifier of the folder need to be updated.</param>
        /// <returns>Update folder request instance that will send to server.</returns>
        protected UpdateFolderType GetUpdateFolderRequest(string folderType, string updateType, FolderIdType folderId)
        {
            UpdateFolderType updateFolderRequest = this.GetUpdateFolderRequest(new string[] { folderType }, new string[] { updateType }, new FolderIdType[] { folderId });
            return updateFolderRequest;
        }

        /// <summary>
        /// Switch from current user to a different user logon to mail box.
        /// </summary>
        /// <param name="logonUser">Name of the user that is about to logon to mail box.</param>
        /// <param name="password">Password of the user that is about to logon to mail box.</param>
        /// <param name="domainValue">Domain of the user that is about to logon to mail box..</param>
        protected void SwitchUser(string logonUser, string password, string domainValue)
        {
            this.FOLDAdapter.SwitchUser(logonUser, password, domainValue);
            this.COREAdapter.SwitchUser(logonUser, password, domainValue);
        }

        /// <summary>
        /// Initialize folder permission related variables.
        /// </summary>
        protected void InitialPermissionVariables()
        {
            this.CanCreateSubFolder = false;
            this.CanReadSubFolder = false;
            this.CanEditSubFolder = false;
            this.CanDeleteSubFolder = false;
            this.CanCreateItem = false;
            this.CanReadOwnedItem = false;
            this.CanEditOwnedItem = false;
            this.CanDeleteOwnedItem = false;
            this.CanReadNotOwnedItem = false;
            this.CanEditNotOwnedItem = false;
            this.CanDeleteNotOwnedItem = false;
        }

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
            impersonation.ConnectingSID.Item = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Set time zone value.
            TimeZoneDefinitionType timezoneDefin = new TimeZoneDefinitionType();
            timezoneDefin.Id = "Eastern Standard Time";
            TimeZoneContextType timezoneContext = new TimeZoneContextType();
            timezoneContext.TimeZoneDefinition = timezoneDefin;

            Dictionary<string, object> headerValues = new Dictionary<string, object>();
            headerValues.Add("MailboxCulture", mailboxCulture);
            headerValues.Add("ExchangeImpersonation", impersonation);
            headerValues.Add("TimeZoneContext", timezoneContext);
            this.FOLDAdapter.ConfigureSOAPHeader(headerValues);
        }

        /// <summary>
        /// Validate if User has related permissions according to the permission level set on folder.
        /// </summary>
        /// <param name="permissionLevel">Permission level value.</param>
        protected void ValidateFolderPermissionLevel(PermissionLevelType permissionLevel)
        {
            #region Create a folder in the User1's inbox folder, and set permission level value for User2

            // Configure permission set.
            PermissionSetType permissionSet = new PermissionSetType();
            permissionSet.Permissions = new PermissionType[1];
            permissionSet.Permissions[0] = new PermissionType();
            permissionSet.Permissions[0].PermissionLevel = permissionLevel;
            permissionSet.Permissions[0].UserId = new UserIdType();
            permissionSet.Permissions[0].UserId.PrimarySmtpAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // CreateFolder request.
            CreateFolderType createFolderRequest = this.GetCreateFolderRequest(DistinguishedFolderIdNameType.inbox.ToString(), new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, new PermissionSetType[] { permissionSet });

            // Create a new folder.
            CreateFolderResponseType createFolderResponse = this.FOLDAdapter.CreateFolder(createFolderRequest);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, createFolderResponse.ResponseMessages.Items[0].ResponseClass, "Fold should be created successfully!");

            // Save the new created folder's folder id.
            FolderIdType newFolderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
            this.NewCreatedFolderIds.Add(newFolderId);

            #endregion

            #region Create an item in the folder created in step 1 with User1's credential

            string itemNameNotOwned = Common.GenerateResourceName(this.Site, "Test Mail");

            // Create an item in the new created folder.
            ItemIdType itemIdNotOwned = this.CreateItem(Common.GetConfigurationPropertyValue("User1Name", Site) + "@" + Common.GetConfigurationPropertyValue("Domain", Site), newFolderId.Id, itemNameNotOwned);
            Site.Assert.IsNotNull(itemIdNotOwned, "Item should be created successfully!");

            #endregion

            #region Switch to User2

            this.SwitchUser(Common.GetConfigurationPropertyValue("User2Name", this.Site), Common.GetConfigurationPropertyValue("User2Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion

            #region Create a subfolder under the folder created in step 1 with User2's credential

            // CreateFolder request.
            CreateFolderType createFolderInSharedMailboxRequest = this.GetCreateFolderRequest(newFolderId.Id, new string[] { "Custom Folder" }, new string[] { "IPF.MyCustomFolderClass" }, null);

            // Create a new folder.
            CreateFolderResponseType createFolderInSharedMailboxResponse = this.FOLDAdapter.CreateFolder(createFolderInSharedMailboxRequest);

            this.CanCreateSubFolder = ResponseClassType.Success == createFolderInSharedMailboxResponse.ResponseMessages.Items[0].ResponseClass;

            #endregion

            #region Edit items that User2 doesn't own with User2's credential

            this.CanEditNotOwnedItem = this.UpdateItemSubject(itemIdNotOwned);
            this.CanReadNotOwnedItem = this.GetItem(itemIdNotOwned);
            this.CanDeleteNotOwnedItem = this.DeleteItem(itemIdNotOwned);

            #endregion

            #region Edit items that User2 owns with User2's credential

            string itemNameOwned = Common.GenerateResourceName(this.Site, "Test Mail");
            string user1MailBox = Common.GetConfigurationPropertyValue("User1Name", Site) + "@" + Common.GetConfigurationPropertyValue("Domain", Site);

            ItemIdType itemIdOwned = this.CreateItem(user1MailBox, newFolderId.Id, itemNameOwned);

            // If user can create items.
            this.CanCreateItem = itemIdOwned != null;
            if (this.CanCreateItem)
            {
                this.CanEditOwnedItem = this.UpdateItemSubject(itemIdOwned);
                this.CanReadOwnedItem = this.GetItem(itemIdOwned);
                this.CanDeleteOwnedItem = this.DeleteItem(itemIdOwned);
            }

            #endregion

            #region Switch to User1

            this.SwitchUser(Common.GetConfigurationPropertyValue("User1Name", this.Site), Common.GetConfigurationPropertyValue("User1Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));

            #endregion
        }

        #endregion
    }
}