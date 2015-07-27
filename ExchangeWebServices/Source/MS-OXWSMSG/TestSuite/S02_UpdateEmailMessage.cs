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
    /// This scenario is designed to test operation related to updating an email message on the server.
    /// </summary>
    [TestClass]
    public class S02_UpdateEmailMessage : TestSuiteBase
    {
        #region Fields
        /// <summary>
        /// The first Item of the first responseMessageItem in infoItems returned from server response.
        /// </summary>
        private ItemType firstItemOfFirstInfoItem;

        /// <summary>
        /// The related Items of ItemInfoResponseMessageType returned from server.
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
        /// This test case is used to verify the related requirements about the server behavior when updating E-mail message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S02_TC01_UpdateMessage()
        {
            #region Create message
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ToRecipients property of the original message
            UpdateItemType updateItemRequest = new UpdateItemType
            {
                MessageDisposition = MessageDispositionType.SaveOnly,
                MessageDispositionSpecified = true,

                ItemChanges = new ItemChangeType[]
                {
                    new ItemChangeType
                    {
                        Item = itemIdType,                        

                        Updates = new ItemChangeDescriptionType[]
                        {
                            new SetItemFieldType
                            {
                                Item = new PathToUnindexedFieldType
                                {
                                    FieldURI = UnindexedFieldURIType.messageToRecipients
                                },

                                // Update the ToRecipients of message from Recipient1 to Recipient2.
                                Item1 = new MessageType
                                {
                                    ToRecipients = new EmailAddressType[]
                                    {
                                        new EmailAddressType
                                        {
                                            EmailAddress = this.Recipient2
                                        }
                                    }
                                }
                            }
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(updateItemResponse), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of message responseMessageItem got from the UpdateItem response.
            if (this.firstItemOfFirstInfoItem != null)
            {
                Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
                itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
                itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            }

            #region Verify the requirements about UpdateItem
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R143");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R143
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                143,
                @"[In UpdateItem] The protocol client sends an UpdateItemSoapIn request WSDL message, and the protocol server responds with an UpdateItemSoapOut response WSDL message.");

            Site.Assert.IsNotNull(this.infoItems[0].ResponseClass, @"The ResponseClass property of the first item of infoItems instance should not be null.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R144");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R144
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                144,
                @"[In UpdateItem] A successful UpdateItem operation request returns an UpdateItemResponse element with the ResponseClass attribute of the UpdateItemResponseMessage element set to ""Success"".");

            Site.Assert.IsNotNull(this.infoItems[0].ResponseCode, @"The ResponseCode property of the first item of infoItems instance should not be null.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R145");
        
            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R145
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                this.infoItems[0].ResponseCode,
                145,
                @"[In UpdateItem] [A successful UpdateItem operation request returns an UpdateItemResponse element] The ResponseCode element of the UpdateItemResponse element is set to ""NoError"".");
            #endregion
            #endregion

            #region Get the updated message
            GetItemType getItemRequest = DefineGeneralGetItemRequestMessage(itemIdType, DefaultShapeNamesType.AllProperties);
            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), @"Server should return success for getting the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(getItemResponse);
            Site.Assert.IsNotNull(this.infoItems, @"The GetItem response should contain one or more items of ItemInfoResponseMessageType.");
            ItemType updatedItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(updatedItem, @"The updated message should exist");

            string expectedValue = Common.GetConfigurationPropertyValue("Recipient2", this.Site);
            Site.Assert.AreEqual<string>(
                expectedValue.ToLower(),
                updatedItem.DisplayTo.ToLower(),
                string.Format("The expected value of the DisplayTo property is {0}. The actual value is {1}.", expectedValue, updatedItem.DisplayTo));
            #endregion

            #region Clean up Sender's drafts folder
            bool isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Sender", this.Site), 
                Common.GetConfigurationPropertyValue("SenderPassword", this.Site), 
                this.Domain, 
                this.Subject, 
                "drafts");
            Site.Assert.IsTrue(isClear, "Sender's drafts folder should be cleaned up.");
            #endregion
        }
        #endregion
    }
}