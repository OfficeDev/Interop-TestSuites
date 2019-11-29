namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving and deletion of an email message with/without optional elements.
    /// </summary>
    [TestClass]
    public class S07_OptionalElementsValidation : TestSuiteBase
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
        /// This test case is used to verify the related requirement about creation, retrieving and deletion of an email message with optional elements.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC01_VerifyMessageWithAllOptionalElements()
        {
            #region Create a message
            #region define a CreateItem request with all elements except ReceivedBy and ReceivedRepresenting elements
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

                // Define the message which contains all the elements except ReceivedBy and ReceivedRepresenting.
                Items = new NonEmptyArrayOfAllItemsType
                {
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                               
                                }
                            },

                            // Specify the recipient that receive a carbon copy of the message.
                            CcRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                    EmailAddress = this.Recipient2
                                }
                            },

                            // Specify the recipient that receive a blind carbon copy of the message.
                            BccRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                    EmailAddress = this.MeetingRoom
                                }
                            },
                            
                            // Specify the subject of message.
                            Subject = this.Subject,

                            // The sender of message does not request a read receipt.
                            IsReadReceiptRequested = false,
                            IsReadReceiptRequestedSpecified = true,

                            // The sender of message does not request a delivery receipt.
                            IsDeliveryReceiptRequested = false,
                            IsDeliveryReceiptRequestedSpecified = true,                           

                            // Response to the message is not requested.
                            IsResponseRequested = false,
                            IsResponseRequestedSpecified = true,

                            // the message has not been read.
                            IsRead = false,
                            IsReadSpecified = true,

                            // Specify the address from whom the message was sent.
                            From = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }
                            },

                            // Specify the address to which replies should be sent.
                            ReplyTo = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                    EmailAddress = this.Recipient2
                                }
                            },
                            
                            // Specify the Usenet header that is used to correlate replies with their original message.
                            References = this.MsgReference,                                                                   
                        }
                    }
                },
            };
            #endregion

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            #endregion

            #region Get the created message via itemIdType in above steps
            GetItemType getItemRequest = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.AllProperties);
            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(getItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), this.infoItems[0].MessageText, null);
            MessageType messageItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0) as MessageType;
            Site.Assert.IsNotNull(messageItem, @"The first item of the array of ItemType type returned from server response should not be null.");

            #region Verify the child elements of the MessageType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R2516, expected the length of EmailAddress is greater than zero, actual length is {0}", messageItem.Sender.Item.EmailAddress.Length);

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R2516
            // The type of the EmailAddress is verified in adapter capture code, only need to check whether this string has a minimum of one character.
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.Sender.Item.EmailAddress.Length > 0,
                2516,
                @"[In t:MessageType Complex Type] When the Mailbox element of Sender element include an EmailAddress element of t:NonEmptyStringType, the t:NonEmptyStringType simple type specifies a string that MUST have a minimum of one character.");

            if (Common.IsRequirementEnabled(1428001, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<MailboxTypeType>(
                MailboxTypeType.Mailbox,
                messageItem.Sender.Item.MailboxType,
                2512,
                @"[In t:MessageType Complex Type] When the Mailbox element of Sender element include an MailboxType element of t:MailboxTypeType, the value ""Mailbox"" of t:MailboxTypeType specifies a mail-enabled directory service object.");
            }

            Site.Assert.AreEqual<string>(this.Sender.ToLower(), messageItem.Sender.Item.EmailAddress.ToLower(), "The EmailAddress of the Sender element in GetItem response should be equal with the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R26");

            // The type of the EmailAddress is verified in adapter capture code, only need to check whether this string has a minimum of one character.
            this.Site.CaptureRequirement(
                26,
                @"[In t:MessageType Complex Type] The Sender element Specifies the sender of a message.");

            Site.Assert.AreEqual<string>(this.Sender.ToLower(), messageItem.From.Item.EmailAddress.ToLower(), "The EmailAddress of the From element in GetItem response should be equal to the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R53");

            // Each user has his own EmailAddress, if the above assert is passed, the following requirements can be captured directly. 
            this.Site.CaptureRequirement(
                53,
                @"[In t:MessageType Complex Type] From element Specifies the addressee from whom the message was sent.");
            
            // The messageRequest is used to save the message in request to create.
            MessageType messageRequest = createItemRequest.Items.Items[0] as MessageType;
            Site.Assert.IsNotNull(messageRequest, @"The CreateItem request should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R28");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R28
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the sender value from getItem response is equal to the setting Sender element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.Sender.Item.EmailAddress.ToLower(),
                messageItem.Sender.Item.EmailAddress.ToLower(),
                28,
                @"[In t:MessageType Complex Type] [Sender element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R55");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R55
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the From value from getItem response is equal to the setting From element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.From.Item.EmailAddress.ToLower(),
                messageItem.From.Item.EmailAddress.ToLower(),
                55,
                @"[In t:MessageType Complex Type] [From element] This is a read/write element.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R4200");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R4200
            this.Site.CaptureRequirementIfIsFalse(
                messageItem.IsReadReceiptRequested,
                4200,
                @"[In t:MessageType Complex Type] [IsReadReceiptRequested element] A text value of ""false"" indicates that a read receipt is not requested from the recipient of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R4600");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R4600
            this.Site.CaptureRequirementIfIsFalse(
                messageItem.IsDeliveryReceiptRequested,
                4600,
                @"[In t:MessageType Complex Type] [IsDeliveryReceiptRequested element] A text value of ""false"" indicates that a delivery receipt has not been requested from the recipient of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R6100");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R6100
            this.Site.CaptureRequirementIfIsFalse(
                messageItem.IsRead,
                6100,
                @"[In t:MessageType Complex Type] [IsRead element]The text value of ""false"" indicates that the message has not been read.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R6500");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R6500
            this.Site.CaptureRequirementIfIsFalse(
                messageItem.IsResponseRequested,
                6500,
                @"[In t:MessageType Complex Type] [IsResponseRequested element] A text value of ""false"" indicates that a response has not been requested.");
            #endregion
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

        /// <summary>
        /// This test case is used to verify the related requirement about creation, retrieving and deletion of an email message without optional elements.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC02_VerifyMessageWithoutOptionalElements()
        {
            #region Create the message
            #region define a CreateItem request without any optional elements of MessageType
            CreateItemType createItemRequestWithoutAllOptionalElements = new CreateItemType
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

                // Define the message which does not contain optional elements.
                Items = new NonEmptyArrayOfAllItemsType
                {
                    Items = new MessageType[]
                    {
                        new MessageType
                        {
                            // Specify the subject of message.
                            Subject = this.Subject,
                        }
                    }
                },
            };
            #endregion

            CreateItemResponseType createItemResponseWithoutAllOptionalElement = this.MSGAdapter.CreateItem(createItemRequestWithoutAllOptionalElements);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponseWithoutAllOptionalElement, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponseWithoutAllOptionalElement);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            #endregion

            #region Get the created message
            GetItemType getItemRequestWithoutAllOptionalElement = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.AllProperties);
            GetItemResponseType getItemResponseWithoutAllOptionalElement = this.MSGAdapter.GetItem(getItemRequestWithoutAllOptionalElement);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponseWithoutAllOptionalElement), @"Server should return success for getting the email messages.");
            Site.Assert.IsNotNull(this.infoItems[0], "The first item of infoItems object should not be null");
            Site.Assert.IsNotNull(this.infoItems[0].ResponseClass, "The ResponseClass property of infoItems[0] object should not be null");
            Site.Assert.IsNotNull(this.infoItems[0].Items, "The Items property of the first item should not be null.");
            #endregion

            #region Verify the optional elements
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R27");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R27
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                27,
                @"[In t:MessageType Complex Type] [Sender element]This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R34");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R34
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                34,
                @"[In t:MessageType Complex Type] [CcRecipients element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R37");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R37
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                37,
                @"[In t:MessageType Complex Type] [BccRecipients element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R40");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R40
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                40,
                @"[In t:MessageType Complex Type] [IsReadReceiptRequested element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R48");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R48
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                48,
                @"[In t:MessageType Complex Type] [ConversationIndex element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R51");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R51
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                51,
                @"[In t:MessageType Complex Type] [ConversationTopic element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R54");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R54
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                54,
                @"[In t:MessageType Complex Type] [From element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R57");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R57
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                57,
                @"[In t:MessageType Complex Type] [InternetMessageId element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R63");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R63
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                63,
                @"[In t:MessageType Complex Type] [IsResponseRequested element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R67");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R67
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                67,
                @"[In t:MessageType Complex Type] [References element] This element is optional.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R70");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R70
            // This requirement can be verified since it's successful when creating the message without any optional properties specified in MessageType.
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                70,
                @"[In t:MessageType Complex Type] [ReplyTo element] This element is optional.");
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

        /// <summary>
        /// This test case is used to verify  related requirement about the ConversationIndex is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC03_VerifyConversationIndexIsReadOnly()
        {
            #region Create a message which include the ConversationIndex element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                               
                                }
                            },
                            ConversationIndex = new byte[] { },                                                                 
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);

            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ConversationIndex is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
           
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createItemResponse.ResponseMessages.Items[0].ResponseClass,
                113001,
                @"[In CreateItem] If the CreateItem WSDL operation is not successful, it returns a CreateItemResponse element with the ResponseClass attribute of the CreateItemResponseMessage element set to ""Error"". ");

            this.Site.CaptureRequirementIfIsTrue(
                System.Enum.IsDefined(typeof(ResponseCodeType), createItemResponse.ResponseMessages.Items[0].ResponseCode),
                113002,
                @"[In CreateItem] [A unsuccessful CreateItem operation request returns a CreateItemResponse element] The ResponseCode element of the CreateItemResponseMessage element is set to one of the common errors defined in [MS-OXWSCDATA] section 2.2.5.24.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ConversationIndex property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageConversationIndex
                                },
                                Item1 = new MessageType
                                {
                                    ConversationIndex = new byte[] { }
                                }
                            }
                        }                   
                    }
                }
            };
  
            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ConversationIndex is read-only, then the The ErrorInvalidPropertySet should be returned by server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R145001");

            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                updateItemResponse.ResponseMessages.Items[0].ResponseClass,
                145001,
                @"[In UpdateItem] If the UpdateItem WSDL operation request is not successful, it returns an UpdateItemResponse element with the ResponseClass attribute of the UpdateItemResponseMessage element set to ""Error"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R145002");

            this.Site.CaptureRequirementIfIsTrue(
                System.Enum.IsDefined(typeof(ResponseCodeType), updateItemResponse.ResponseMessages.Items[0].ResponseCode),
                145002,
                @"[In UpdateItem] [A unsuccessful UpdateItem operation request returns an UpdateItemResponse element] The ResponseCode element of the UpdateItemResponseMessage element is set to one of the common errors defined in [MS-OXWSCDATA] section 2.2.5.24.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R49");

            // If the above steps are pass, the R49 will be verified.
            this.Site.CaptureRequirement(
                49,
                @"[In t:MessageType Complex Type] [ConversationIndex element] This element is read-only.");
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

        /// <summary>
        /// This test case is used to verify  related requirement about the ConversationTopic is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC04_VerifyConversationTopicIsReadOnly()
        {
            #region Create a message which include the ConversationTopic element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                               
                                }
                            },
                            ConversationTopic = "test"                                                               
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);

            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ConversationTopic is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ConversationTopic property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageConversationTopic
                                },
                                Item1 = new MessageType
                                {
                                     ConversationTopic = "test"      
                                }
                            }
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ConversationTopic is read-only, then the The ErrorInvalidPropertySet should be returned by server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R52");

            // If the above steps are pass, the R52 will be verified.
            this.Site.CaptureRequirement(
                52,
                @"[In t:MessageType Complex Type] [ConversationTopic element] This element is read-only.");
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

        /// <summary>
        /// This test case is used to verify  related requirement about the InternetMessageId is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC05_VerifyInternetMessageIdIsReadOnly()
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

            #region update InternetMessageId property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageInternetMessageId
                                },
                                Item1 = new MessageType
                                {
                                     InternetMessageId = "InternetMessageId"     
                                }
                            }
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since InternetMessageId is read-only, then the The ErrorInvalidPropertySet should be returned by server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R58");

            // If the above steps are pass, the R58 will be verified.
            this.Site.CaptureRequirement(
                58,
                @"[In t:MessageType Complex Type] [InternetMessageId element] This element is read-only.");
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

        /// <summary>
        /// This test case is used to verify  related requirement about the ReceivedBy is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC06_VerifyReceivedByIsReadOnly()
        {
            #region Create a message which include the ReceivedBy element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                               
                                }
                            },
                            ReceivedBy = new SingleRecipientType()
                            {
                                Item = new EmailAddressType()
                                {
                                    EmailAddress = this.Recipient1
                                }
                            }
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);

            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ConversationIndex is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update InternetMessageId property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageReceivedBy
                                },
                                Item1 = new MessageType
                                {
                                     ReceivedBy = new SingleRecipientType()
                                    {
                                        Item = new EmailAddressType()
                                        {
                                            EmailAddress = this.Recipient1
                                        }
                                    }    
                                }
                            }
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedBy is read-only, then the The ErrorInvalidPropertySet should be returned by server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R72001");

            // If the above steps are pass, the R72001 will be verified.
            this.Site.CaptureRequirement(
                72001,
                @"[In t:MessageType Complex Type] [ReceivedBy element] This element is read-only.");
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

        /// <summary>
        /// This test case is used to verify  related requirement about the ReceivedRepresenting is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC07_VerifyReceivedRepresentingIsReadOnly()
        {
            #region Create a message which include the ReceivedRepresenting element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }                                
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1                               
                                }
                            },
                           ReceivedRepresenting = new SingleRecipientType()
                           {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Recipient1
                                }
                           }
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);

            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ReceivedRepresenting property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageReceivedRepresenting
                                },
                                Item1 = new MessageType
                                {
                                     ReceivedRepresenting = new SingleRecipientType()
                                     {
                                         Item = new EmailAddressType()
                                         {
                                             EmailAddress = this.Recipient1
                                         }
                                     }
                                }
                            }
                        }                   
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73001");

            // If the above steps are pass, the R73001 will be verified.
            this.Site.CaptureRequirement(
                73001,
                @"[In t:MessageType Complex Type] [ReceivedRepresenting element] This element is read-only.");

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

        /// <summary>
        /// This test case is used to verify  related requirement about the ApprovalRequestData is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC08_VerifyApprovalRequestDataIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(73008, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the ApprovalRequestData element.");
            #region Create a message which include the ApprovalRequestData element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the ApprovalRequestData of the message.
                            ApprovalRequestData = new ApprovalRequestDataType()
                            {
                                IsUndecidedApprovalRequest=false
                            }
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ReceivedRepresenting property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageReceivedRepresenting
                                },
                                Item1 = new MessageType
                                {
                                    // Specify the ApprovalRequestData of the message.
                                    ApprovalRequestData = new ApprovalRequestDataType()
                                    {
                                        IsUndecidedApprovalRequest=false
                                    }
                                }
                            }
                        }
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73008");

            // If the above steps are pass, the R73008 will be verified.
            this.Site.CaptureRequirement(
                73008,
                @"[In t:MessageType Complex Type] [ApprovalRequestData] This element is read-only.");

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

        /// <summary>
        /// This test case is used to verify  related requirement about the VotingInformation is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC09_VerifyVotingInformationIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(73009, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the VotingInformation element.");
            #region Create a message which include the VotingInformation element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the VotingInformation of the message.
                            VotingInformation = new VotingInformationType()
                            {
                                UserOptions=new VotingOptionDataType[1]
                            }
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ReceivedRepresenting property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageReceivedRepresenting
                                },
                                Item1 = new MessageType
                                {
                                    // Specify the VotingInformation of the message.
                                    VotingInformation = new VotingInformationType()
                                    {
                                        UserOptions=new VotingOptionDataType[1]
                                    }
                                }
                            }
                        }
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73009");

            // If the above steps are pass, the R73009 will be verified.
            this.Site.CaptureRequirement(
                73009,
                @"[In t:MessageType Complex Type] [VotingInformation]This element is read-only.");

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

        /// <summary>
        /// This test case is used to verify  related requirement about the ReminderMessageData is read-only. 
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S07_TC10_VerifyReminderMessageDataIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(73010, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the ReminderMessageData element.");
            #region Create a message which include the ReminderMessageData element.
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
                    Items = new MessageType[]
                    {
                        // Create a MessageType instance with all element.
                        new MessageType
                        {
                            // Specify the sender of the message.
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.Sender
                                }
                            },

                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the ReminderMessageData of the message.
                            ReminderMessageData = new ReminderMessageDataType()
                            {
                                ReminderText="ReminderText"
                            }
                        }
                    }
                },
            };

            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, createItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            #region Create message
            createItemRequest = this.GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Save the ItemId of message responseMessageItem returned from the createItem response.
            ItemIdType itemIdType = new ItemIdType();
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
            itemIdType.Id = this.firstItemOfFirstInfoItem.ItemId.Id;
            itemIdType.ChangeKey = this.firstItemOfFirstInfoItem.ItemId.ChangeKey;
            #endregion

            #region update ReceivedRepresenting property of the original message
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
                                    FieldURI = UnindexedFieldURIType.messageReceivedRepresenting
                                },
                                Item1 = new MessageType
                                {
                                    // Specify the ReminderMessageData of the message.
                                    ReminderMessageData = new ReminderMessageDataType()
                                    {
                                        ReminderText="ReminderText"
                                    }
                                }
                            }
                        }
                    }
                }
            };

            UpdateItemResponseType updateItemResponse = this.MSGAdapter.UpdateItem(updateItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorInvalidPropertySet, updateItemResponse.ResponseMessages.Items[0].ResponseCode, "Since ReceivedRepresenting is read-only, then the The ErrorInvalidPropertySet should be returned by server.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73010");

            // If the above steps are pass, the R73010 will be verified.
            this.Site.CaptureRequirement(
                73010,
                @"[In t:MessageType Complex Type] [ReminderMessageData]This element is read-only.");

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