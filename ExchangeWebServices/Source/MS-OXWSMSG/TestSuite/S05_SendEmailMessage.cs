namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operation related to sending of an email message on the server.
    /// </summary>
    [TestClass]
    public class S05_SendEmailMessage : TestSuiteBase
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
        /// This test case is used to verify the related requirements about the server behavior when sending E-mail message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S05_TC01_SendMessage()
        {
            #region Create the message
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem got from the createItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            ItemIdType itemIdType = new ItemIdType();
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region Send the message
            SendItemType sendItemRequest = new SendItemType
            {
                ItemIds = new ItemIdType[]
                {
                    itemIdType
                },

                SaveItemToFolder = true,

                // Save the message copy in sent items folder.
                SavedItemFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.sentitems
                    }
                }
            };

            SendItemResponseType sendItemResponse = this.MSGAdapter.SendItem(sendItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(sendItemResponse), @"Server should return success for sending the email messages.");
            Site.Assert.IsNotNull(sendItemResponse.ResponseMessages.Items, @"Items in the returned response should not be null.");

            #region Verify the requirements about SendItem
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R177");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R177           
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                177,
                @"[In SendItem] The protocol client sends a SendItemSoapIn request WSDL message, and the protocol server responds with a SendItemSoapOut response WSDL message.");

            Site.Assert.IsNotNull(sendItemResponse.ResponseMessages.Items[0].ResponseClass, @"The ResponseClass property of the first item of infoItems instance should not be null.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R178");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R178
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                sendItemResponse.ResponseMessages.Items[0].ResponseClass,
                178,
                @"[In SendItem] A successful SendItem operation request returns a SendItemResponse element with the ResponseClass attribute of the SendItemResponseMessage element set to ""Success"".");

            Site.Assert.IsNotNull(sendItemResponse.ResponseMessages.Items[0].ResponseCode, @"The ResponseCode property of the first item of infoItems instance should not be null.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R179");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R179
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                sendItemResponse.ResponseMessages.Items[0].ResponseCode,
                179,
                @"[In SendItem] [A successful SendItem operation request returns a SendItemResponse element] The ResponseCode element of the SendItemResponse element is set to ""NoError"".");
            #endregion
            #endregion

            #region Create the invalid message for SendItem operation without ToRecipients element
            CreateItemType createItemRequestFailed = new CreateItemType
            {
                MessageDisposition = MessageDispositionType.SaveOnly,

                // MessageDispositionSpecified value needs to be set.
                MessageDispositionSpecified = true,

                SavedItemFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.drafts
                    }
                },

                Items = new NonEmptyArrayOfAllItemsType
                {
                    Items = new MessageType[]
                    {
                        new MessageType
                        {
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },
                            Subject = this.Subject,                                          
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponseFailed = this.MSGAdapter.CreateItem(createItemRequestFailed);
            Site.Assert.IsTrue(this.VerifyResponse(createItemResponseFailed), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponseFailed);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of message responseMessageItem got from the createItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region Send the message without ToRecipients element
            SendItemType sendItemRequestFailed = new SendItemType
            {
                // Set to invalid message's ItemId.
                ItemIds = new ItemIdType[]
                {
                    itemIdType
                },

                SaveItemToFolder = true,
                SavedItemFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.sentitems
                    }
                }
            };

            SendItemResponseType sendItemResponseFailed = this.MSGAdapter.SendItem(sendItemRequestFailed);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R31");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R31   
            Site.CaptureRequirementIfAreNotEqual<ResponseClassType>(
                ResponseClassType.Success,
                sendItemResponseFailed.ResponseMessages.Items[0].ResponseClass,
                31,
                @"[In t:MessageType Complex Type] [ToRecipients element] This element is required for sending a message.");
            #endregion

            #region Delete the invalid message
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                // Set to invalid message's ItemId.
                ItemIds = new ItemIdType[]
                {
                    itemIdType
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");
            #endregion

            #region Clean up Sender's sentitems folder and Recipient1's inbox folder
            bool isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Sender", this.Site), 
                Common.GetConfigurationPropertyValue("SenderPassword", this.Site), 
                this.Domain, 
                this.Subject, 
                "sentitems");
            Site.Assert.IsTrue(isClear, "Sender's sentitems folder should be cleaned up.");

            isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Recipient1", this.Site), 
                Common.GetConfigurationPropertyValue("Recipient1Password", this.Site), 
                this.Domain, 
                this.Subject, 
                "inbox");
            Site.Assert.IsTrue(isClear, "Recipient1's inbox folder should be cleaned up.");
            #endregion
        }
        #endregion 
    }
}