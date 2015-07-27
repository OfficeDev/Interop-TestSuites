//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to movement of the contact items in the server.
    /// </summary>
    [TestClass]
    public class S04_MoveContactItem : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S04_TC01_MoveContactItem()
        {
            #region Step 1:Create the contact item.
            // Create a contact item.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);
            #endregion

            #region Step 2:Move the contact item.
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

            #region Step 3:Get the moved contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
            #endregion

            #region Step 4:Get the original contact item Id.
            // Call GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            Site.Assert.AreEqual<int>(
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0));

            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorItemNotFound,
                getItemResponse.ResponseMessages.Items[0].ResponseCode,
                string.Format(
                    "Get contact item with original item Id should fail! Expected response code: {0}, actual response code: {1}",
                    ResponseCodeType.ErrorItemNotFound,
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion
        }
        #endregion
    }
}