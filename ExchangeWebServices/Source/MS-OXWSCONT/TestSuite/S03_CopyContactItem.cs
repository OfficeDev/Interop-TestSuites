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
    /// This scenario is designed to test operation related to copy of the contact items in the server.
    /// </summary>
    [TestClass]
    public class S03_CopyContactItem : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S03_TC01_CopyContactItem()
        {
            #region Step 1:Create the contact item.
            // Create a contact item.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2:Copy the contact item.
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

            #region Step 3:Get the contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);
            #endregion
        }
        #endregion
    }
}