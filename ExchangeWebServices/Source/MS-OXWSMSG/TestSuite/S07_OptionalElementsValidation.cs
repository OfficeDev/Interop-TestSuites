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
        #endregion
    }
}