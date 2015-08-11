namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, updating, movement, retrieving, copy and deletion of the multiple contact items in the server.
    /// </summary>
    [TestClass]
    public class S05_OperateMultipleContactItems : TestSuiteBase
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
        /// This test case is intended to validate the successful response of operating multiple contact items, returned by CreateItem, UpdateItem, MoveItem, GetItem and CopyItem operations.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S05_TC01_OperateMultipleContactItems()
        {
            #region Step 1:Create the contact item.
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the contact items
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ContactItemType[2];

            // Create the first contact item.
            createItemRequest.Items.Items[0] = this.BuildContactItemWithRequiredProperties();

            // Create the second contact item.
            createItemRequest.Items.Items[1] = this.BuildContactItemWithRequiredProperties();

            // Configure the SavedItemFolderId of CreateItem request to specify that the created item is saved under which folder.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType()
            {
                Item = new DistinguishedFolderIdType()
                {
                    Id = DistinguishedFolderIdNameType.contacts,
                }
            };
            #endregion

            CreateItemResponseType createItemResponse = this.CONTAdapter.CreateItem(createItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 2, this.Site);
            #endregion

            #region Step 2:Update the contact items.
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
                    },
                    new ItemChangeType()
                    {
                        Item = this.ExistContactItems[1],

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
                                    FileAs = FileAsMappingType.LastCommaFirst.ToString()
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
            Common.CheckOperationSuccess(updateItemResponse, 2, this.Site);
            #endregion

            #region Step 3:Move the contact items.
            MoveItemType moveItemRequest = new MoveItemType();
            MoveItemResponseType moveItemResponse = new MoveItemResponseType();

            // Configure ItemIds.
            moveItemRequest.ItemIds = new BaseItemIdType[2];
            moveItemRequest.ItemIds[0] = this.ExistContactItems[0];
            moveItemRequest.ItemIds[1] = this.ExistContactItems[1];

            // Clear existContactItems for MoveItem.
            this.InitializeCollection();

            // Configure move Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.drafts;
            moveItemRequest.ToFolderId = new TargetFolderIdType();
            moveItemRequest.ToFolderId.Item = distinguishedFolderId;

            moveItemResponse = this.CONTAdapter.MoveItem(moveItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(moveItemResponse, 2, this.Site);
            #endregion

            #region Step 4:Get the contact items.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);
            #endregion

            #region Step 5:Copy the contact items.
            CopyItemType copyItemRequest = new CopyItemType();
            CopyItemResponseType copyItemResponse = new CopyItemResponseType();

            // Configure ItemIds.
            copyItemRequest.ItemIds = new BaseItemIdType[2];
            copyItemRequest.ItemIds[0] = this.ExistContactItems[0];
            copyItemRequest.ItemIds[1] = this.ExistContactItems[1];

            // Configure the copy Distinguished Folder.
            DistinguishedFolderIdType distinguishedFolderIdForCopyItem = new DistinguishedFolderIdType();
            distinguishedFolderIdForCopyItem.Id = DistinguishedFolderIdNameType.drafts;
            copyItemRequest.ToFolderId = new TargetFolderIdType();
            copyItemRequest.ToFolderId.Item = distinguishedFolderIdForCopyItem;

            copyItemResponse = this.CONTAdapter.CopyItem(copyItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(copyItemResponse, 2, this.Site);
            #endregion
        }
        #endregion
    }
}