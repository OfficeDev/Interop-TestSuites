namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, updating, movement, retrieving and copy of the contact items with optional elements in the server.
    /// </summary>
    [TestClass]
    public class S06_OperateContactItemWithOptionalElements : TestSuiteBase
    {
        #region Class initialize and clean up
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
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
        /// This test case is intended to validate the successful response of operating contact item without optional elements, returned by CreateItem, UpdateItem, MoveItem, GetItem and CopyItem operations.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S06_TC01_OperateContactItemWithoutOptionalElements()
        {
            #region Step 1:Create the contact item.
            // Create a contact item without optional elements.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2:Update the contact item.
            UpdateItemType updateItemRequest = new UpdateItemType()
            {
                // Configure ItemIds.
                ItemChanges = new ItemChangeType[]
                {
                    new ItemChangeType()
                    {
                        Item = this.ExistContactItems[0],

                        Updates = new ItemChangeDescriptionType[]
                        {
                            new SetItemFieldType()
                            {
                                Item = new PathToUnindexedFieldType()
                                {
                                    FieldURI = UnindexedFieldURIType.contactsFileAs
                                },

                                Item1 = new ContactItemType()
                                {
                                    FileAs = FileAsMappingType.LastFirstCompany.ToString()
                                }
                            }
                        }
                    }
                },

                ConflictResolution = ConflictResolutionType.AlwaysOverwrite
            };

            // Clear existContactItems for MoveItem.
            this.InitializeCollection();

            UpdateItemResponseType updateItemResponse = new UpdateItemResponseType();

            // Invoke UpdateItem operation.
            updateItemResponse = this.CONTAdapter.UpdateItem(updateItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);
            #endregion

            #region Step 3:Move the contact item.
            MoveItemType moveItemRequest = new MoveItemType();
            MoveItemResponseType moveItemResponse = new MoveItemResponseType();

            // Configure ItemIds.
            moveItemRequest.ItemIds = new BaseItemIdType[1];
            moveItemRequest.ItemIds[0] = this.ExistContactItems[0];

            // Clear existContactItems for MoveItem.
            this.InitializeCollection();

            // Configure move Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.drafts;
            moveItemRequest.ToFolderId = new TargetFolderIdType();
            moveItemRequest.ToFolderId.Item = distinguishedFolderId;

            moveItemResponse = this.CONTAdapter.MoveItem(moveItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            #endregion

            #region Step 4:Get the contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
            #endregion

            #region Step 5:Copy the contact item.
            CopyItemType copyItemRequest = new CopyItemType();
            CopyItemResponseType copyItemResponse = new CopyItemResponseType();

            // Configure ItemIds.
            copyItemRequest.ItemIds = new BaseItemIdType[1];
            copyItemRequest.ItemIds[0] = this.ExistContactItems[0];

            // Configure the copy Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderIdForCopyItem = new DistinguishedFolderIdType();
            distinguishedFolderIdForCopyItem.Id = DistinguishedFolderIdNameType.drafts;
            copyItemRequest.ToFolderId = new TargetFolderIdType();
            copyItemRequest.ToFolderId.Item = distinguishedFolderIdForCopyItem;

            copyItemResponse = this.CONTAdapter.CopyItem(copyItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response of operating contact item with optional elements, returned by CreateItem, UpdateItem, MoveItem, GetItem and CopyItem operations.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S06_TC02_OperateContactItemWithOptionalElements()
        {
            #region Step 1:Create the contact item.
            // Create a full property contact item.
            ContactItemType item = this.CreateFullPropertiesContact();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2:Update the contact item.
            UpdateItemType updateItemRequest = new UpdateItemType()
            {
                // Configure ItemIds.
                ItemChanges = new ItemChangeType[]
                {
                    new ItemChangeType()
                    {
                        Item = this.ExistContactItems[0],

                        Updates = new ItemChangeDescriptionType[]
                        {
                            new SetItemFieldType()
                            {
                                Item = new PathToUnindexedFieldType()
                                {
                                    FieldURI = UnindexedFieldURIType.contactsFileAs
                                },

                                Item1 = new ContactItemType()
                                {
                                    FileAs = FileAsMappingType.LastFirstCompany.ToString()
                                }
                            }
                        }
                    }
                },

                ConflictResolution = ConflictResolutionType.AlwaysOverwrite
            };

            // Clear existContactItems for MoveItem.
            this.InitializeCollection();

            UpdateItemResponseType updateItemResponse = new UpdateItemResponseType();

            // Invoke UpdateItem operation.
            updateItemResponse = this.CONTAdapter.UpdateItem(updateItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);
            #endregion

            #region Step 3:Move the contact item.
            MoveItemType moveItemRequest = new MoveItemType();
            MoveItemResponseType moveItemResponse = new MoveItemResponseType();

            // Configure ItemIds.
            moveItemRequest.ItemIds = new BaseItemIdType[1];
            moveItemRequest.ItemIds[0] = this.ExistContactItems[0];

            // Clear existContactItems for MoveItem.
            this.InitializeCollection();

            // Configure move Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.drafts;
            moveItemRequest.ToFolderId = new TargetFolderIdType();
            moveItemRequest.ToFolderId.Item = distinguishedFolderId;

            moveItemResponse = this.CONTAdapter.MoveItem(moveItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            #endregion

            #region Step 4:Get the contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);
 
            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
            #endregion

            #region Step 5:Copy the contact item.
            CopyItemType copyItemRequest = new CopyItemType();
            CopyItemResponseType copyItemResponse = new CopyItemResponseType();

            // Configure ItemIds.
            copyItemRequest.ItemIds = new BaseItemIdType[1];
            copyItemRequest.ItemIds[0] = this.ExistContactItems[0];

            // Configure the copy Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderIdForCopyItem = new DistinguishedFolderIdType();
            distinguishedFolderIdForCopyItem.Id = DistinguishedFolderIdNameType.drafts;
            copyItemRequest.ToFolderId = new TargetFolderIdType();
            copyItemRequest.ToFolderId.Item = distinguishedFolderIdForCopyItem;

            copyItemResponse = this.CONTAdapter.CopyItem(copyItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);
            #endregion
        }
        #endregion
    }
}