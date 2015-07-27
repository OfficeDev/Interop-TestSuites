//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to movement of an email message on the server.
    /// </summary>
    [TestClass]
    public class S04_MoveEmailMessage : TestSuiteBase
    {
        #region Fields
        /// <summary>
        /// The first Item of the first responseMessageItem in infoItems returned from server response.
        /// </summary>
        private ItemType firstItemOfFirstInfoItem;

        /// <summary>
        /// The related Item of ItemInfoResponseMessageType returned from server.
        /// </summary>
        private ItemInfoResponseMessageType[] infoItems;
        #endregion

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
        /// This test case is used to verify the related requirements about the server behavior when moving E-mail message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S04_TC01_MoveMessage()
        {
            #region Create message
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem got from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region Move message
            MoveItemType moveItemRequest = new MoveItemType
            {
                ItemIds = new ItemIdType[]
                {
                    itemIdType
                },

                // Set target folder to junk email folder.
                ToFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.junkemail
                    }
                }
            };

            MoveItemResponseType moveItemResponse = this.MSGAdapter.MoveItem(moveItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(moveItemResponse), @"Server should return success for moving the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(moveItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of message responseMessageItem got from the moveItem response.
            ItemIdType moveItemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            moveItemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            moveItemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;

            // Verify whether the message is moved to junkemail folder.
            string userName = Common.GetConfigurationPropertyValue("Sender", this.Site);
            string password = Common.GetConfigurationPropertyValue("SenderPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            bool findItemInDrafts = this.SRCHSUTControlAdapter.IsItemAvailableAfterMoveOrDelete(userName, password, domain, "drafts", this.Subject, "itemSubject");
            Site.Assert.IsFalse(findItemInDrafts, "The item should not be found in the drafts folder of Sender.");
            
            bool findItemInJunkemail = this.SearchItems(Role.Sender, "junkemail", this.Subject, "itemSubject");
            Site.Assert.IsTrue(findItemInJunkemail, "The item should be found in the junkemail folder of Sender.");

            #region Verify the requirements about MoveItem
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R161");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R161            
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                161,
                @"[In MoveItem] The protocol client sends a MoveItemSoapIn request WSDL message, and the protocol server responds with a MoveItemSoapOut response WSDL message.");

            Site.Assert.IsNotNull(this.infoItems[0].ResponseClass, @"The ResponseClass property of the first item of infoItems instance should not be null.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R162");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R162
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                162,
                @"[In MoveItem] A successful MoveItem operation request returns a MoveItemResponse element with the ResponseClass attribute of the MoveItemResponseMessage element set to ""Success"".");

            Site.Assert.IsNotNull(this.infoItems[0].ResponseCode, @"The ResponseCode property of the first item of infoItems instance should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R163");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R163
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                this.infoItems[0].ResponseCode,
                163,
                @"[In MoveItem] [A successful MoveItem operation request returns a MoveItemResponse element] The ResponseCode element of the MoveItemResponse element is set to ""NoError"".");
            #endregion
            #endregion

            #region Delete the moved message
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   moveItemIdType
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");
            #endregion
        }
        #endregion
    }
}