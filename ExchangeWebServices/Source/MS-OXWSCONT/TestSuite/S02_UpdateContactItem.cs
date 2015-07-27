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
    /// This scenario is designed to test operation related to updating of the contact items in the server.
    /// </summary>
    [TestClass]
    public class S02_UpdateContactItem : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S02_TC01_UpdateContactItem()
        {
            #region Step 1:Create the contact item.
            // Create a contact item.
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

            UpdateItemResponseType updateItemResponse = new UpdateItemResponseType();

            // Invoke UpdateItem operation.
            updateItemResponse = this.CONTAdapter.UpdateItem(updateItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);
            #endregion

            #region Step 3:Get the contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ContactItemType[] contacts = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

            Site.Assert.AreEqual<int>(
                1,
                contacts.Length,
                string.Format(
                    "The count of items from response should be 1, actual: '{0}'.", contacts.Length));

            Site.Assert.AreEqual<string>(
                FileAsMappingType.LastFirstCompany.ToString(),
                contacts[0].FileAs,
                string.Format(
                    "The FileAs property should be updated as set. Expected value: {0}, actual value: {1}", FileAsMappingType.LastFirstCompany.ToString(), contacts[0].FileAs));
            #endregion
        }
        #endregion
    }
}