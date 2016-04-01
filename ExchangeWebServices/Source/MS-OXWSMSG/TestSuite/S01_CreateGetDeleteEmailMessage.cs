namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, and deletion of an email message from the server.
    /// </summary>
    [TestClass]
    public class S01_CreateGetDeleteEmailMessage : TestSuiteBase
    {
        #region Fields
        /// <summary>
        /// The related Items of ItemInfoResponseMessageType returned from server.
        /// </summary>
        private ItemInfoResponseMessageType[] infoItems;

        /// <summary>
        /// The first Item of the first responseMessageItem in infoItems returned from server response.
        /// </summary>
        private ItemType firstItemOfFirstInfoItem;
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
        /// This test case is designed to test successfully creating, getting and deleting a single message after creating the single calendar item with all properties (except about Delegate) of MessageType.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC01_CreateGetDeleteMessageWithAllElementsExceptAboutDelegate()
        {
            #region Create a message with all elements except ReceivedBy and ReceivedRepresenting elements
            #region Define a CreateItem request
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
                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the recipient that receives a carbon copy of the message.
                            CcRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                    EmailAddress = this.Recipient2
                                }
                            },

                            // Specify the recipient that receives a blind carbon copy of the message.
                            BccRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                    EmailAddress = this.MeetingRoom
                                }
                            },
                            
                            // Specify the subject of message.
                            Subject = this.Subject,

                            // The sender of message requests a read receipt.
                            IsReadReceiptRequested = true,
                            IsReadReceiptRequestedSpecified = true,

                            // The sender of message requests a delivery receipt.
                            IsDeliveryReceiptRequested = true,
                            IsDeliveryReceiptRequestedSpecified = true,                           

                            // Response to the message is requested.
                            IsResponseRequested = true,
                            IsResponseRequestedSpecified = true,

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
                            IsRead = true,
                            IsReadSpecified = true,
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

            #region Verify the requirements about CreateItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R111");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R111            
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                111,
                @"[In CreateItem] The protocol client sends a CreateItemSoapIn request WSDL message, and the protocol server responds with a CreateItemSoapOut response WSDL message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R112");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R112
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                112,
                @"[In CreateItem]If the CreateItem WSDL operation request is successful, the server returns a CreateItemResponse element, as specified in [MS-OXWSCORE] section 3.1.4.2.2.2, with the ResponseClass attribute, as specified in [MS-OXWSCDATA] section 2.2.4.67, of the CreateItemResponseMessage element, as specified in [MS-OXWSCDATA] section 2.2.4.12, set to ""Success"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R113");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R113
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                this.infoItems[0].ResponseCode,
                113,
                @"[In CreateItem] [A successful CreateItem operation request returns a CreateItemResponse element] The ResponseCode element, as specified in [MS-OXWSCDATA] section 2.2.4.67, of the CreateItemResponseMessage element is set to ""NoError"". ");
            #endregion

            #region Get the created message via itemIdType in above steps
            #region The getItemResponseOfFirstMessage is used to store the getItemResponseOfFirstMessage gotten from server when the BaseShape element is set to IdOnly.
            GetItemType getItemRequest = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.IdOnly);
            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), "The response of the GetItem should be valid.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R127");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R127
            bool isVerifyR127 = Common.IsIdOnly((XmlElement)this.MSGAdapter.LastRawResponseXml, "t:Message", "t:ItemId");
            Site.CaptureRequirementIfIsTrue(
                isVerifyR127,
                127,
                @"[In GetItem] [The BaseShape child element in ItemShape element] Set this element[BaseShape element] to IdOnly to return only the item ID.");
            #endregion

            #region The getItemResponseOfFirstMessage is used to store the getItemResponseOfFirstMessage got from server when the BaseShape element is set to AllProperties.
            getItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), @"Server should return success for getting the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(getItemResponse);
            Site.Assert.IsNotNull(this.infoItems, @"The GetItem response should contain one or more items of ItemInfoResponseMessageType.");
            MessageType messageItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0) as MessageType;
            Site.Assert.IsNotNull(messageItem, @"The message item returned from server response should not be null.");

            #region Verify requirements about GetItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R133");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R133            
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                133,
                @"[In GetItem] The protocol client sends a GetItemSoapIn request WSDL message, and the protocol server responds with a GetItemSoapOut response WSDL message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R134");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R134
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                this.infoItems[0].ResponseClass,
                134,
                @"[In GetItem] If the GetItem WSDL operation request is successful, the server returns a GetItemResponse element, as specified in [MS-OXWSCORE] section 3.1.4.4.2.2, with the ResponseClass attribute, as specified in [MS-OXWSCDATA] section 2.2.4.67, of the GetItemResponseMessage element, as specified in [MS-OXWSCDATA] section 2.2.4.12, set to ""Success"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R135");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R135
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                this.infoItems[0].ResponseCode,
                135,
                @"[In GetItem] [A successful GetItem operation request returns a GetItemResponse element] The ResponseCode element, as specified in [MS-OXWSCDATA] section 2.2.4.67, of the GetItemResponseMessage element is set to ""NoError"". ");

            bool isVerifiedR198 = Common.IsIdOnly((XmlElement)this.MSGAdapter.LastRawResponseXml, "t:Message", "t:ItemId");

            this.Site.CaptureRequirementIfIsFalse(
                isVerifiedR198,
                198, 
                @"[In GetItem] [The BaseShape child element in ItemShape element] Set it[BaseShape element] to AllProperties to return all of the properties used by the server to construct a message.");
            #endregion

            #region Verify the child elements of the MessageType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R2916, expected the length of EmailAddress is greater than zero, actual length is {0}", messageItem.ToRecipients[0].EmailAddress.Length);

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R2916
            // The type of the EmailAddress is verified in adapter capture code, only need to check whether this string has a minimum of one character.
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.ToRecipients[0].EmailAddress.Length > 0,
                2916,
                @"[In t:MessageType Complex Type] When the Mailbox element of ToRecipients element include an EmailAddress element of t:NonEmptyStringType, the t:NonEmptyStringType simple type specifies a string that MUST have a minimum of one character.");

            if (Common.IsRequirementEnabled(1428001, this.Site))
            {
                this.Site.CaptureRequirementIfAreEqual<MailboxTypeType>(
                    MailboxTypeType.Mailbox,
                    messageItem.ToRecipients[0].MailboxType,
                    2912,
                    @"[In t:MessageType Complex Type]When the Mailbox element of ToRecipients element include an MailboxType element of t:MailboxTypeType, the value ""Mailbox"" of t:MailboxTypeType specifies a mail-enabled directory service object.");
            }

            Site.Assert.AreEqual<string>(this.Recipient1.ToLower(), messageItem.ToRecipients[0].EmailAddress.ToLower(), "The EmailAddress of the ToRecipients element in GetItem response should be equal with the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R30");

            // Each user has his own EmailAddress, if the above assert is passed, the following requirements can be captured directly. 
            this.Site.CaptureRequirement(
                30,
                @"[In t:MessageType Complex Type] ToRecipients element Specifies a collection of recipients of an email.");

            Site.Assert.AreEqual<string>(this.Recipient2.ToLower(), messageItem.CcRecipients[0].EmailAddress.ToLower(), "The EmailAddress of the CcRecipients element in GetItem response should be equal with the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R33");

            // Each user has his own EmailAddress, if the above assert is passed, the following requirements can be captured directly.
            this.Site.CaptureRequirement(
                33,
                @"[In t:MessageType Complex Type] CcRecipients element Specifies a collection of recipients that receive a carbon copy (Cc) of an email.");

            Site.Assert.AreEqual<string>(this.MeetingRoom.ToLower(), messageItem.BccRecipients[0].EmailAddress.ToLower(), "The EmailAddress of the BccRecipients element in GetItem response should be equal with the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R36");

            // Each user has his own EmailAddress, if the above assert is passed, the following requirements can be captured directly.
            this.Site.CaptureRequirement(
                36,
                @"[In t:MessageType Complex Type] BccRecipients element Specifies a collection of recipients that receive a blind carbon copy (Bcc) of an email.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R42");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R42
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.IsReadReceiptRequested,
                42,
                @"[In t:MessageType Complex Type] [IsReadReceiptRequested element] A text value of ""true"" indicates that a read receipt is requested from the recipient of the message.");

            // If the R42 is verified, then the IsReadReceiptRequested is a boolean and it indicates that a read receipt is requested from the recipient of the message.
            // So the R39 will be verified.
            this.Site.CaptureRequirement(
                39,
                @"[In t:MessageType Complex Type] IsReadReceiptRequested element Specifies a Boolean value that indicates whether the sender of a message requests a read receipt.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R46");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R46            
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.IsDeliveryReceiptRequested,
                46,
                @"[In t:MessageType Complex Type] [IsDeliveryReceiptRequested element] A text value of ""true"" indicates that a delivery receipt has been requested from the recipient of the message.");

            // If the R46 is verified, then the IsDeliveryReceiptRequested is a boolean and it indicates that a delivery receipt has been requested from the recipient of the message.
            // So the R43 will be verified.
            this.Site.CaptureRequirement(
                43,
                @"[In t:MessageType Complex Type] IsDeliveryReceiptRequested element Specifies a Boolean value that indicates whether the sender of the message has requested a delivery receipt.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R61");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R61            
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.IsRead,
                61,
                @"[In t:MessageType Complex Type] [IsRead element]The text value of ""true"" indicates that the message has been read.");

            // If R61 is verified, then the IsRead is boolean and it indicates that the message has been read.
            // So the R59 will be verified.
            this.Site.CaptureRequirement(
                59,
                @"[In t:MessageType Complex Type] IsRead element Specifies a Boolean value that indicates whether the message has been read.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R65");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R65            
            this.Site.CaptureRequirementIfIsTrue(
                messageItem.IsResponseRequested,
                65,
                @"[In t:MessageType Complex Type] [IsResponseRequested element] A text value of ""true"" indicates that a response has been requested.");

            // If R65 is verified, then the IsResponseRequested is boolean and it indicates that a response has been requested.
            // So R62 will be verified.
            this.Site.CaptureRequirement(
                62,
                @"[In t:MessageType Complex Type] IsResponseRequested element Specifies a Boolean value that indicates whether a response to an email has been requested.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R66");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R66
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.MsgReference,
                messageItem.References,
                66,
                @"[In t:MessageType Complex Type] References element Specifies the Usenet header that is used to correlate replies with their original message.");

            Site.Assert.AreEqual<string>(this.Recipient2.ToLower(), messageItem.ReplyTo[0].EmailAddress.ToLower(), "The EmailAddress of the From element in GetItem response should be equal to the settings");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R69");

            // Each user has his own EmailAddress, if the above assert is passed, the following requirements can be captured directly.
            this.Site.CaptureRequirement(
                69,
                @"[In t:MessageType Complex Type] ReplyTo element Specifies a collection of addresses to send replies to. ");
            #endregion

            #region Verify the requirements about the read/write element of MessageType
            // The messageRequest is used to save the message in request to create.
            MessageType messageRequest = createItemRequest.Items.Items[0] as MessageType;
            Site.Assert.IsNotNull(messageRequest, @"The CreateItem request should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R32");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R32
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the ToRecipients value from getItem response is equal to the setting ToRecipients element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.ToRecipients[0].EmailAddress.ToLower(),
                messageItem.ToRecipients[0].EmailAddress.ToLower(),
                32,
                @"[In t:MessageType Complex Type] [ToRecipients element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R35");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R35
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the CcRecipients value from getItem response is equal to the setting CcRecipients element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.CcRecipients[0].EmailAddress.ToLower(),
                messageItem.CcRecipients[0].EmailAddress.ToLower(),
                35,
                @"[In t:MessageType Complex Type] [CcRecipients element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R38");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R38
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the BccRecipients value from getItem response is equal to the setting BccRecipients element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.BccRecipients[0].EmailAddress.ToLower(),
                messageItem.BccRecipients[0].EmailAddress.ToLower(),
                38,
                @"[In t:MessageType Complex Type] [BccRecipients element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R41");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R41
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the IsReadReceiptRequested value from getItem response is equal to the setting IsReadReceiptRequested element value in createItem request.
            Site.CaptureRequirementIfAreEqual<bool>(
                messageRequest.IsReadReceiptRequested,
                messageItem.IsReadReceiptRequested,
                41,
                @"[In t:MessageType Complex Type] [IsReadReceiptRequested element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R45");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R45
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the IsDeliveryReceiptRequested value from getItem response is equal to the setting IsDeliveryReceiptRequested element value in createItem request.
            Site.CaptureRequirementIfAreEqual<bool>(
                messageRequest.IsDeliveryReceiptRequested,
                messageItem.IsDeliveryReceiptRequested,
                45,
                @"[In t:MessageType Complex Type] [IsDeliveryReceiptRequested element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R64");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R64
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the IsResponseRequested value from getItem response is equal to the setting IsResponseRequested element value in createItem request.
            Site.CaptureRequirementIfAreEqual<bool>(
                messageRequest.IsResponseRequested,
                messageItem.IsResponseRequested,
                64,
                @"[In t:MessageType Complex Type] [IsResponseRequested element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R68");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R68
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the References value from getItem response is equal to the setting References element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.References,
                messageItem.References,
                68,
                @"[In t:MessageType Complex Type] [References element] This is a read/write element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R71");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R71
            // This requirement can be verified since the createItem response message succeeded which indicates message was created successfully 
            // and the ReplyTo value from getItem response is equal to the setting ReplyTo element value in createItem request.
            Site.CaptureRequirementIfAreEqual<string>(
                messageRequest.ReplyTo[0].EmailAddress.ToLower(),
                messageItem.ReplyTo[0].EmailAddress.ToLower(),
                71,
                @"[In t:MessageType Complex Type] [ReplyTo element] This is a read/write element.");

            this.Site.CaptureRequirementIfAreEqual<bool>(
                messageRequest.IsRead,
                messageItem.IsRead,
                60,
                @"[In t:MessageType Complex Type] [IsRead element] This is a read/write element.");
            #endregion
            #endregion
            #endregion

            #region Delete the message created
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   this.firstItemOfFirstInfoItem.ItemId
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");

            GetItemType getdeletedItemRequest = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.IdOnly);
            GetItemResponseType getdeletedItemResponse = this.MSGAdapter.GetItem(getdeletedItemRequest);

            // verify whether the creating message deleted
            if (getdeletedItemResponse != null
                && getdeletedItemResponse.ResponseMessages != null
                && getdeletedItemResponse.ResponseMessages.Items != null)
            {
                if (getdeletedItemResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
                {
                    Site.Assert.Fail("Server should not return success for get the email messages information, because the created message have been deleted.");
                }
            }

            #region Verify requirements about DeleteItem operation
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R152");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R152            
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                152,
                @"[In DeleteItem] The protocol client sends a DeleteItemSoapIn request WSDL message, and the protocol server responds with a DeleteItemSoapOut response WSDL message.");

            Site.Assert.IsNotNull(deleteItemResponse.ResponseMessages.Items[0], @"The first item of the array of InfoItems in server response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R153");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R153
            Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                deleteItemResponse.ResponseMessages.Items[0].ResponseClass,
                153,
                @"[In DeleteItem] If the DeleteItem WSDL operation request is successful, the server returns a DeleteItemResponse element, as specified in [MS-OXWSCORE] section 3.1.4.3.2.2, with the ResponseClass attribute, as specified in [MS-OXWSCDATA] section 2.2.4.67, of the DeleteItemResponseMessage element, as specified in [MS-OXWSCDATA] section 2.2.4.12, set to ""Success"". ");

            Site.Assert.IsNotNull(deleteItemResponse.ResponseMessages.Items[0].ResponseCode, @"The ResponseCode property of the first item of the array of InfoItems should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R154");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R154
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                deleteItemResponse.ResponseMessages.Items[0].ResponseCode,
                154,
                @"[In DeleteItem] [A successful DeleteItem operation request returns a DeleteItemResponse element] The ResponseCode element, as specified by [MS-OXWSCDATA] section 2.2.4.67, of the DeleteItemResponseMessage element is set to ""NoError"". ");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the related requirements about SaveOnly value is set in MessageDispositionType when create a new message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC02_CreateMessageWithMessageDispositionTypeSaveOnly()
        {
            #region Sender creates a message
            CreateItemType createItemRequestOfMessage = GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.inbox);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequestOfMessage);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R82");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R82
            // This requirement can be verified when the ItemId of message responseMessageItem got from the createItem response is not null which indicates the responseMessageItem identifier is returned.
            Site.CaptureRequirementIfIsNotNull(
                this.firstItemOfFirstInfoItem.ItemId,
                82,
                @"[In t:MessageDispositionType Simple Type] [when ""SaveOnly"" used in CreateItem Type element] In this case, an item identifier is returned.");
            #endregion

            #region Sender gets the message
            GetItemType getItemRequest = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.AllProperties);
            GetItemResponseType getItemResponseOfMessage = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponseOfMessage), @"Server should return success for getting the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(getItemResponseOfMessage);
            Site.Assert.IsNotNull(this.infoItems, @"The CreateItem response should contain one or more items of ItemInfoResponseMessageType.");
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");
            #endregion

            #region Sender finds the message in the inbox folder
            bool findItemResult = this.SearchItems(Role.Sender, "inbox", this.Subject, "itemSubject");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R80");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R80
            // This requirement can be verified since the message created using SaveOnly can be sent successfully by using the SendItem operation.
            Site.CaptureRequirementIfIsTrue(
                findItemResult,
                80,
                @"[In t:MessageDispositionType Simple Type] The value ""SaveOnly"" means when used in the CreateItemType complex type ([MS-OXWSCORE] section 3.1.4.2.3.2), the email message item is saved in the folder that is specified by the TargetFolderIdType complex type ([MS-OXWSFOLD] section 2.2.4.16).");
            #endregion

            #region Sender sends the created message
            SendItemType sendItemRequest = new SendItemType
            {
                ItemIds = new ItemIdType[] { this.firstItemOfFirstInfoItem.ItemId },

                SaveItemToFolder = true,
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
            #endregion

            #region Recipient1 verifies if the message has been received
            findItemResult = this.SearchItems(Role.Recipient1, "inbox", this.Subject, "itemSubject");
            Site.Assert.IsTrue(findItemResult, "The item should be found in the inbox folder of Recipient1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R81");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R81
            // This requirement can be verified since the message created using SaveOnly can be sent successfully by using the SendItem operation.
            Site.CaptureRequirementIfIsTrue(
                findItemResult,
                81,
                @"[In t:MessageDispositionType Simple Type] [when ""SaveOnly"" used in CreateItem Type element] Messages can be sent later by using the SendItem operation (section 3.1.4.7) on an ExchangeServiceBinding object.");
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

        /// <summary>
        /// This test case is used to verify the related requirements about SendOnly value is set in MessageDispositionType when create a new message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC03_CreateMessageWithMessageDispositionTypeSendOnly()
        {
            #region Sender creates a message
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SendOnly, DistinguishedFolderIdNameType.junkemail);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SendOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R84");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R84
            // This requirement can be verified when the responseMessageItem is null which indicates there is no responseMessageItem identifier being returned from server.
            Site.CaptureRequirementIfIsNull(
                this.infoItems[0].Items.Items,
                84,
                @"[In t:MessageDispositionType Simple Type] [When ""SendOnly"" used in CreateItem Type element] In this case, an item identifier is not returned.");
            #endregion

            #region Sender verifies if the message is in the junkmail folder
            bool findItemResult = true;
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            while (findItemResult && upperBound-- > 0)
            {               
                findItemResult = this.SearchItems(Role.Sender, "junkemail", this.Subject, "itemSubject");
                System.Threading.Thread.Sleep(waitTime);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R203");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R203
            // This requirement can be verified after the message send out, there is not a copy saved in the specified folder.
            Site.CaptureRequirementIfIsFalse(
                findItemResult,
                203,
                @"[In t:MessageDispositionType Simple Type] The value ""SendOnly"" means when used in the CreateItemType complex type, the email message item [is sent] but no copy is saved.");
            #endregion

            #region Recipient1 verifies if the message has been received
            findItemResult = this.SearchItems(Role.Recipient1, "inbox", this.Subject, "itemSubject");
            Site.Assert.IsTrue(findItemResult, "The item should be found in the inbox folder of Recipient1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R83");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R83
            Site.CaptureRequirementIfIsTrue(
                findItemResult,
                83,
                @"[In t:MessageDispositionType Simple Type] The value ""SendOnly"" means when used in the CreateItemType complex type, the email message item is sent.");
            #endregion

            #region Recipient1 find the the message that has been received
            string userName = Common.GetConfigurationPropertyValue("Recipient1", this.Site);
            string password = Common.GetConfigurationPropertyValue("Recipient1Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            MessageType receivedMessage = this.FindSpecificItem(userName, password, domain, "inbox", this.Subject, "itemSubject") as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R72");

            this.Site.CaptureRequirementIfAreEqual<string>(
                userName,
                receivedMessage.ReceivedBy.Item.Name,
                72,
                @"[In t:MessageType Complex Type] ReceivedBy element Identifies the delegate in a delegate access scenario.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R196");

            this.Site.CaptureRequirementIfIsInstanceOfType(
                receivedMessage.ReceivedBy,
                typeof(SingleRecipientType),
                196,
                @"[In t:MessageType Complex Type] ReceivedBy element is t:SingleRecipientType type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73");

            this.Site.CaptureRequirementIfAreEqual<string>(
                userName,
                receivedMessage.ReceivedRepresenting.Item.Name,
                73,
                @"[In t:MessageType Complex Type] ReceivedRepresenting element Identifies the principal in a delegate access scenario.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R197");

            this.Site.CaptureRequirementIfIsInstanceOfType(
                receivedMessage.ReceivedRepresenting,
                typeof(SingleRecipientType),
                197,
                @"[In t:MessageType Complex Type] ReceivedRepresenting element is t:SingleRecipientType");
            #endregion

            #region Clean up Recipient1's inbox folder
            bool isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Recipient1", this.Site),
                Common.GetConfigurationPropertyValue("Recipient1Password", this.Site),
                this.Domain,
                this.Subject,
                "inbox");
            Site.Assert.IsTrue(isClear, "Recipient1's inbox folder should be cleaned up.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the related requirements about SendAndSaveCopy value is set in MessageDispositionType when create a new message.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC04_CreateMessageWithMessageDispositionTypeSendAndSaveCopy()
        {
            #region Sender creates message with MessageDisposition setting to SendAndSaveCopy.
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SendAndSaveCopy, DistinguishedFolderIdNameType.inbox);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SendAndSaveCopy), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R86");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R86
            // This requirement can be verified when the responseMessageItem is null which indicates there is no responseMessageItem identifier is returned from server.
            Site.CaptureRequirementIfIsNull(
                this.infoItems[0].Items.Items,
                86,
                @"[In t:MessageDispositionType Simple Type] [When ""SendAndSaveCopy"" used in CreateItem Type element] In this case, an item identifier is not returned.");
            #endregion

            #region Sender verifies if the message has been saved in the inbox folder
            bool findItemResult = this.SearchItems(Role.Sender, "inbox", this.Subject, "itemSubject");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R204");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R204
            // This requirement can be verified when the responseMessageItem can be found in saved Folder, which is specified in CreateItem request. The response is not null means the message is found.
            Site.CaptureRequirementIfIsTrue(
                findItemResult,
                204,
                @"[In t:MessageDispositionType Simple Type] The value ""SendAndSaveCopy"" means when used in the CreateItemType complex type, the email message item [is sent] and a copy is saved in the TargetFolderIdType complex type.");
            #endregion

            #region Recipient1 verifies if the message has been received
            findItemResult = this.SearchItems(Role.Recipient1, "inbox", this.Subject, "itemSubject");
            Site.Assert.IsTrue(findItemResult, "The item should be found in the inbox folder of Recipient1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R85");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R85
            // This requirement can be verified since the ResponseClass of CreateItemResponse is "success", which indicates the e-mail message created is sent successfully.
            Site.CaptureRequirementIfIsTrue(
                findItemResult,
                85,
                @"[In t:MessageDispositionType Simple Type] The value ""SendAndSaveCopy"" means when used in the CreateItemType complex type, the email message item is sent.");
            #endregion

            #region Clean up Sender's inbox folder and Recipient1's inbox folder
            bool isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Sender", this.Site),
                Common.GetConfigurationPropertyValue("SenderPassword", this.Site),
                this.Domain,
                this.Subject,
                "inbox");
            Site.Assert.IsTrue(isClear, "Sender's inbox folder should be cleaned up.");

            isClear = this.MSGSUTControlAdapter.CleanupFolders(
                Common.GetConfigurationPropertyValue("Recipient1", this.Site),
                Common.GetConfigurationPropertyValue("Recipient1Password", this.Site),
                this.Domain,
                this.Subject,
                "inbox");
            Site.Assert.IsTrue(isClear, "Recipient1's inbox folder should be cleaned up.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify the related requirements about call GetItem unsuccessful.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC05_GetAndDeleteMessageUnsuccessful()
        {
            #region Create a message with all elements except ReceivedBy and ReceivedRepresenting elements
            CreateItemType createItemRequest = GetCreateItemType(MessageDispositionType.SaveOnly, DistinguishedFolderIdNameType.drafts);
            CreateItemResponseType createItemResponse = this.MSGAdapter.CreateItem(createItemRequest);
            Site.Assert.IsTrue(this.VerifyCreateItemResponse(createItemResponse, MessageDispositionType.SaveOnly), @"Server should return success for creating the email messages.");
            this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
            this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
            #endregion

            #region Delete the message created
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   this.firstItemOfFirstInfoItem.ItemId
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");
            #endregion

            #region Get the created message via itemIdType in above steps
            GetItemType getItemRequest = DefineGeneralGetItemRequestMessage(this.firstItemOfFirstInfoItem.ItemId, DefaultShapeNamesType.IdOnly);
            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R135001");

            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                getItemResponse.ResponseMessages.Items[0].ResponseClass,
                135001,
                @"[In GetItem] If the GetItem WSDL operation request is not successful, it returns a GetItemResponse element with the ResponseClass attribute of the GetItemResponseMessage element set to ""Error"". ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R135002");

            this.Site.CaptureRequirementIfIsTrue(
                System.Enum.IsDefined(typeof(ResponseCodeType), getItemResponse.ResponseMessages.Items[0].ResponseCode),
                135002,
                @"[In GetItem] [A unsuccessful GetItem operation request returns a GetItemResponse element] The ResponseCode element of the GetItemResponseMessage element is set to one of the common errors defined in [MS-OXWSCDATA] section 2.2.5.24.");
            #endregion

            #region Delete the message deleted.
            ItemIdType itemId = new ItemIdType()
            {
                Id = this.firstItemOfFirstInfoItem.ItemId.Id + "invalid",
                ChangeKey = this.firstItemOfFirstInfoItem.ItemId.Id,
            };

            deleteItemRequest.ItemIds = new ItemIdType[] { itemId };
            deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                deleteItemResponse.ResponseMessages.Items[0].ResponseClass,
                154001,
                @"[In DeleteItem] If the DeleteItem WSDL operation request is not successful, it returns a DeleteItemResponse element with the ResponseClass attribute of the DeleteItemResponseMessage element set to ""Error"". ");
            
            this.Site.CaptureRequirementIfIsTrue(
                System.Enum.IsDefined(typeof(ResponseCodeType), deleteItemResponse.ResponseMessages.Items[0].ResponseCode),
                154002,
                @"[In DeleteItem] [A unsuccessful DeleteItem operation request returns a DeleteItemResponse element] The ResponseCode element of the DeleteItemResponseMessage element is set to one of the common errors defined in [MS-OXWSCDATA] section 2.2.5.24.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to test reqiurements related with the IncludeMimeContent element.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC06_GetMessageWithIncludeMimeContent()
        {
            #region Create a message
            #region Define a CreateItem request
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
                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the subject of message.
                            Subject = this.Subject,
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
            #region Send the GetItem requst which the IncludeMimeContent is set true.
            GetItemType getItemRequest = new GetItemType
            {
                ItemIds = new ItemIdType[]
                {
                    this.firstItemOfFirstInfoItem.ItemId
                },

                ItemShape = new ItemResponseShapeType
                {
                    BaseShape = DefaultShapeNamesType.AllProperties,
                    IncludeMimeContent = true,
                    IncludeMimeContentSpecified = true,
                }
            };
           
            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), "The response of the GetItem should be valid.");
            MessageType message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21188");

            this.Site.CaptureRequirementIfIsNotNull(
                message.MimeContent,
                "MS-OXWSCDATA",
                21188,
                @"[In t:ItemResponseShapeType Complex Type] [IncludeMimeContent is] True, specifies the MIME content of an item is returned in a response.");
            #endregion

            #region Send the GetItem requst which the IncludeMimeContent is set false.
            getItemRequest.ItemShape.IncludeMimeContent = false;
            getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), @"Server should return success for getting the email messages.");
            message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21189");

            this.Site.CaptureRequirementIfIsNull(
               message.MimeContent,
               "MS-OXWSCDATA",
               21189,
               @"[In t:ItemResponseShapeType Complex Type] otherwise [IncludeMimeContent is] false, specifies [the MIME content of an item is not returned in a response].");
            #endregion
            #endregion

            #region Delete the message created
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   this.firstItemOfFirstInfoItem.ItemId
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");

            #endregion
        }

        /// <summary>
        /// This test case is designed to test reqiurements related with the FilterHtmlContent element.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC07_GetMessageWithFilterHtmlContent()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("SutVersion", this.Site).Equals("ExchangeServer2007") == false, "Exchange 2007 does not use the FilterHtmlContent element.");

            #region Create a message with html body
            #region Define a CreateItem request
            string unsafeHtmlContent = "<html><body><b>Bold</b><script>alert('Alert!');</script></body></html>";

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
                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the subject of message.
                            Subject = this.Subject,

                           Body = new BodyType()
                           {
                               BodyType1 = BodyTypeType.HTML,
                               Value = unsafeHtmlContent,
                           }
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
            #region Send the GetItem requst which the FilterHtmlContent is set true.
            GetItemType getItemRequest = new GetItemType
            {
                ItemIds = new ItemIdType[]
                {
                    this.firstItemOfFirstInfoItem.ItemId
                },

                ItemShape = new ItemResponseShapeType
                {
                    BaseShape = DefaultShapeNamesType.AllProperties,
                    FilterHtmlContent = true,
                    FilterHtmlContentSpecified = true,
                }
            };

            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), "The response of the GetItem should be valid.");
            MessageType message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21193");

            this.Site.CaptureRequirementIfIsFalse(
                 message.Body.Value.Contains("<script>alert('Alert!');</script>"),
                 "MS-OXWSCDATA",
                 21193,
                 @"[In t:ItemResponseShapeType Complex Type] [FilterHtmlContent is] True, specifies HTML content filtering is enabled.");
            #endregion

            #region Send the GetItem requst which the FilterHtmlContent is set false.
            getItemRequest.ItemShape.FilterHtmlContent = false;
            getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), @"Server should return success for getting the email messages.");
            message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21194");

            this.Site.CaptureRequirementIfIsTrue(
                 message.Body.Value.Contains("<script>alert('Alert!');</script>"),
                 "MS-OXWSCDATA",
                 21194,
                 @"[In t:ItemResponseShapeType Complex Type] otherwise [FilterHtmlContent is] false, specifies [HTML content filtering is not enabled].");
            #endregion
            #endregion

            #region Delete the message created
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   this.firstItemOfFirstInfoItem.ItemId
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");

            #endregion
        }

        /// <summary>
        /// This test case is designed to test reqiurements related with the BodyType element and AdditionalProperties element.
        /// </summary>
        [TestCategory("MSOXWSMSG"), TestMethod()]
        public void MSOXWSMSG_S01_TC08_GetMessageWithBodyTypeAndAdditionProperties()
        {
            #region Create a message.
            #region Define a CreateItem request
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
                            // Specify the recipient of the message.
                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                     EmailAddress = this.Recipient1
                                }
                            },

                            // Specify the subject of message.
                            Subject = this.Subject,
                            Body = new BodyType()
                            {
                                BodyType1 = BodyTypeType.HTML,
                                Value = "<html><body><b>Bold</b>test</body></html>"
                            }
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
            #region Send the GetItem requst which the BodyType element is set to text.
            GetItemType getItemRequest = new GetItemType
            {
                ItemIds = new ItemIdType[]
                {
                    this.firstItemOfFirstInfoItem.ItemId
                },

                ItemShape = new ItemResponseShapeType
                {
                    BaseShape = DefaultShapeNamesType.AllProperties,
                    BodyType = BodyTypeResponseType.Text,
                    BodyTypeSpecified = true,
                }
            };

            GetItemResponseType getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), "The response of the GetItem should be valid.");
            MessageType message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1190");

            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                 BodyTypeType.Text,
                 message.Body.BodyType1,
                 "MS-OXWSCDATA",
                 1190,
                 @"[In t:ItemResponseShapeType Complex Type] The element ""BodyType"" with type ""t:BodyTypeResponseType(section 2.2.3.1)"" specifies the requested body text format for the Body property that is returned in a response.");
            #endregion

            #region Send the GetItem requst which the BaseShape element is set to IdOnly and addition the subject property.
            getItemRequest.ItemShape.BaseShape = DefaultShapeNamesType.IdOnly;
            getItemRequest.ItemShape.AdditionalProperties = new BasePathToElementType[] 
            {
                new PathToUnindexedFieldType()
                {
                    FieldURI = UnindexedFieldURIType.itemSubject,
                }
            };
            getItemResponse = this.MSGAdapter.GetItem(getItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(getItemResponse), @"Server should return success for getting the email messages.");
            message = ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0] as MessageType;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1197");

            this.Site.CaptureRequirementIfIsTrue(
                 string.IsNullOrEmpty(message.Subject) == false,
                 "MS-OXWSCDATA",
                 1197,
                 @"[In t:ItemResponseShapeType Complex Type] The element ""AdditionalProperties"" with type ""t:NonEmptyArrayOfPathsToElementType"" specifies a set of requested additional properties to return in a response.");
            #endregion
            #endregion

            #region Delete the message created
            DeleteItemType deleteItemRequest = new DeleteItemType
            {
                ItemIds = new ItemIdType[]
                {
                   this.firstItemOfFirstInfoItem.ItemId
                }
            };

            DeleteItemResponseType deleteItemResponse = this.MSGAdapter.DeleteItem(deleteItemRequest);
            Site.Assert.IsTrue(this.VerifyResponse(deleteItemResponse), @"Server should return success for deleting the email messages.");

            #endregion
        }
        #endregion
    }
}