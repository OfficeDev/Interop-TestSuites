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
    /// This scenario is designed to test operations related to creation, retrieving, updating, copy, movement, sending and deletion of multiple email messages on the server at the same time.
    /// </summary>
    [TestClass]
    public class S06_OperateMultipleEmailMessages : TestSuiteBase
    {
        #region Fields
        /// <summary>
        /// The first Item of the first responseMessageItem in infoItems returned from server response.
        /// </summary>
        private ItemType firstItemOfFirstInfoItem;

        /// <summary>
        /// The private field specifies the first responseMessageItem of the second responseMessageItem of infoItems returned from server.
        /// </summary>
        private ItemType firstItemOfSecondInfoItem;

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
        /// This test case is used to verify the related requirements about the server behavior when operating multiple E-mail messages at the same time.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S06_TC01_OperateMultipleMessages()
        {
            #region Create multiple message
            string subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("Subject", Site), 0);
            string anotherSubject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("Subject", Site), 1);

            CreateItemType createItemRequest = new CreateItemType
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
                    // Create an responseMessageItem with two MessageType instances.
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

                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                              
                                }
                            },

                            Subject = subject,                                          
                        },

                        new MessageType
                        {
                             Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient2                               
                                }
                            },

                            Subject = anotherSubject,
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(createItemResponse), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the first ItemId of message responseMessageItem got from the createItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            ItemIdType itemIdType1 = new ItemIdType();            
            itemIdType1.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType1.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;

            // Save the second ItemId.
            this.firstItemOfSecondInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 1, 0);
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem, @"The second item of the array of ItemType type returned from server response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem.ItemId, @"The ItemId property of the second item should not be null.");
            ItemIdType itemIdType2 = new ItemIdType();
            itemIdType2.Id = this.firstItemOfSecondInfoItem.ItemId.Id;
            itemIdType2.ChangeKey = this.firstItemOfSecondInfoItem.ItemId.ChangeKey;
            #endregion

            #region Get the multiple messages which created
            GetItemType getItemRequest = new GetItemType
            {
                // Set the two ItemIds got from CreateItem response.
                ItemIds = new ItemIdType[]
                {
                    itemIdType1,
                    itemIdType2
                },

                ItemShape = new ItemResponseShapeType
                {
                    BaseShape = DefaultShapeNamesType.AllProperties,
                }
            };

            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(getItemResponse), @"Server should return success for creating the email messages.");
            #endregion

            #region Update the multiple messages which created
            UpdateItemType updateItemRequest = new UpdateItemType
            {
                MessageDisposition = MessageDispositionType.SaveOnly,

                // MessageDispositionSpecified value needs to be set.
                MessageDispositionSpecified = true,

                // Create two ItemChangeType instances.
                ItemChanges = new ItemChangeType[]
                {
                    new ItemChangeType
                    {
                        Item = itemIdType1,                        

                        Updates = new ItemChangeDescriptionType[]
                        {
                            new SetItemFieldType
                            {
                                Item = new PathToUnindexedFieldType
                                {
                                    FieldURI = UnindexedFieldURIType.messageToRecipients
                                },

                                // Update ToRecipients property of the first message.
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
                            },                           
                        }                   
                    },

                    new ItemChangeType
                    {
                        Item = itemIdType2,                        

                        Updates = new ItemChangeDescriptionType[]
                        {
                            new SetItemFieldType
                            {
                                Item = new PathToUnindexedFieldType
                                {
                                    FieldURI = UnindexedFieldURIType.messageToRecipients
                                },

                                // Update ToRecipients property of the second message.
                                Item1 = new MessageType
                                {
                                    ToRecipients = new EmailAddressType[]
                                    {
                                        new EmailAddressType
                                        {
                                            EmailAddress = this.Recipient1
                                        }
                                    }
                                }
                            },                           
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(updateItemResponse), @"Server should return success for updating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(updateItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of the first message responseMessageItem got from the UpdateItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType1.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType1.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;

            // Save the ItemId of the second message responseMessageItem got from the UpdateItem response.
            this.firstItemOfSecondInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 1, 0);
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem, @"The second item of the array of ItemType type returned from server response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem.ItemId, @"The ItemId property of the second item should not be null.");
            itemIdType2.Id = this.firstItemOfSecondInfoItem.ItemId.Id;
            itemIdType2.ChangeKey = this.firstItemOfSecondInfoItem.ItemId.ChangeKey;
            #endregion

            #region Copy the updated multiple message to junkemail
            CopyItemType copyItemRequest = new CopyItemType
            {
                ItemIds = new ItemIdType[]
                {
                    itemIdType1,
                    itemIdType2,
                },

                // Copy the message to junk email folder.
                ToFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.junkemail
                    }
                }
            };

            CopyItemResponseType copyItemResponse = this.MSGAdapter.CopyItem(copyItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(copyItemResponse), @"Server should return success for copying the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(copyItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of the first message responseMessageItem got from the CopyItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            ItemIdType copyItemIdType1 = new ItemIdType();
            copyItemIdType1.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            copyItemIdType1.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;

            // Save the ItemId of the second message responseMessageItem got from the CopyItem response.
            this.firstItemOfSecondInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 1, 0);
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem, @"The second item of the array of ItemType type returned from server response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem.ItemId, @"The ItemId property of the second item should not be null.");
            ItemIdType copyItemIdType2 = new ItemIdType();
            copyItemIdType2.Id = this.firstItemOfSecondInfoItem.ItemId.Id;
            copyItemIdType2.ChangeKey = this.firstItemOfSecondInfoItem.ItemId.ChangeKey;
            #endregion

            #region Move the copied multiple message from junkemail to deleteditems
            MoveItemType moveItemRequest = new MoveItemType
            {
                // Set to copied message responseMessageItem id.
                ItemIds = new ItemIdType[]
                {
                    copyItemIdType1,
                    copyItemIdType2
                },

                // Move the copied messages to deleted items folder.
                ToFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = DistinguishedFolderIdNameType.deleteditems
                    }
                }
            };

            MoveItemResponseType moveItemResponse = this.MSGAdapter.MoveItem(moveItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(moveItemResponse), @"Server should return success for moving the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(moveItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.infoItems, @"InfoItems in the returned response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            // Save the ItemId of the first message responseMessageItem got from the MoveItem response.
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            copyItemIdType1.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            copyItemIdType1.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;

            // Save the ItemId of the second message responseMessageItem got from the MoveItem response.
            this.firstItemOfSecondInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 1, 0);
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem, @"The second item of the array of ItemType type returned from server response should not be null.");
            Site.Assert.IsNotNull(this.firstItemOfSecondInfoItem.ItemId, @"The ItemId property of the second item should not be null.");
            copyItemIdType2.Id = this.firstItemOfSecondInfoItem.ItemId.Id;
            copyItemIdType2.ChangeKey = this.firstItemOfSecondInfoItem.ItemId.ChangeKey;
            #endregion

            #region Send multiple messages
            SendItemType sendItemRequest = new SendItemType
            {
                // Set to the two updated messages' ItemIds.
                ItemIds = new ItemIdType[]
                {
                    itemIdType1,
                    itemIdType2
                },

                // Do not save copy.
                SaveItemToFolder = false,
            };

            SendItemResponseType sendItemResponse = this.MSGAdapter.SendItem(sendItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(sendItemResponse), @"Server should return success for sending the email messages.");
            #endregion

            #region Delete the copied messages
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                // Set to the two copied messages' ItemIds.
                ItemIds = new ItemIdType[]
                {
                   copyItemIdType1,
                   copyItemIdType2
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyMultipleResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");
            #endregion

            #region Clean up Recipient1's and Recipient2's inbox folders
            bool isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Recipient2", this.Site), 
                Common.GetConfigurationPropertyValue("Recipient2Password", this.Site), 
                this.Domain, 
                subject, 
                "inbox");
            Site.Assert.IsTrue(isClear, "Recipient2's inbox folder should be cleaned up.");

            isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Recipient1", this.Site), 
                Common.GetConfigurationPropertyValue("Recipient1Password", this.Site), 
                this.Domain, 
                anotherSubject, 
                "inbox");
            Site.Assert.IsTrue(isClear, "Recipient1's inbox folder should be cleaned up.");
            #endregion
        }
        #endregion
    }
}