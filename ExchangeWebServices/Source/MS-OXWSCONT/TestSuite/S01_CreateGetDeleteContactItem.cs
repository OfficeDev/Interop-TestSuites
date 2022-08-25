namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving and deletion of the contact items in the server.
    /// </summary>
    [TestClass]
    public class S01_CreateGetDeleteContactItem : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, GetItem and DeleteItem operations for contact item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC01_VerifyContactItemWithRequiredElement()
        {
            #region Step 1:Create the contact item.
            // Call CreateItem operation.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2:Get the contact item.
            // The contact item to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
            this.ExistContactItems.CopyTo(itemArray, 0);

            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            // Get the item ids from response.
            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            Site.Assert.AreEqual<int>(
                1,
                getItemIds.GetLength(0),
                "One contact item should be returned!");
            #endregion

            #region Step 3:Delete the contact item.
            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.ItemIds = getItemIds;

            // Configure the enumeration value that specifies how an contact item is to be deleted.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            DeleteItemResponseType deleteItemResponse = this.CONTAdapter.DeleteItem(deleteItemRequest);

            // Check the response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 4:Get the deleted contact item
            // Call GetItem operation.
            getItemResponse = this.CallGetItemOperation(getItemIds);

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
                    "Get deleted item should fail! Expected response code: {0}, actual response code: {1}",
                    ResponseCodeType.ErrorItemNotFound,
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem and GetItem operations for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC02_VerifyContactItemWithAllElement()
        {
            #region Step 1:Create the all property contact item.
            // Create a full property contact item.
            ContactItemType item = this.CreateFullPropertiesContact();
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            // Check the response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2:Get the created contact item.
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

            this.VerifyChildElementInContactItemTypeComplexType(contacts[0], item);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the EmailAddressKeyType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC03_VerifyContactItemWithEmailAddressKeyTypeEnums()
        {
            // The value count of enumeration "EmailAddressKeyType" is 3.
            int enumCount = 3;
            EmailAddressKeyType[] emailAddressKeyTypes = new EmailAddressKeyType[enumCount];

            emailAddressKeyTypes[0] = EmailAddressKeyType.EmailAddress1;
            emailAddressKeyTypes[1] = EmailAddressKeyType.EmailAddress2;
            emailAddressKeyTypes[2] = EmailAddressKeyType.EmailAddress3;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain an EmailAddressKeyType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                EmailAddressKeyType emailAddressKeyType = emailAddressKeyTypes[i];

                #region Step 1:Create the contact item with EmailAddressKeyType
                // Create a contact item with EmailAddressKeyType
                ContactItemType item = this.BuildContactItemWithEmailAddress(emailAddressKeyType);
                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item.
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsNotNull(
                    contacts[i].EmailAddresses,
                    "The EmailAddresses element in returned contact item should not be null.");

                Site.Assert.AreEqual<int>(
                    1,
                    contacts[i].EmailAddresses.GetLength(0),
                    "One entry of EmailAddresses element should be returned!");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R124");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R124
            Site.CaptureRequirementIfAreEqual<EmailAddressKeyType>(
                EmailAddressKeyType.EmailAddress1,
                contacts[0].EmailAddresses[0].Key,
                124,
                @"[In t:EmailAddressKeyType Simple Type] EmailAddress1: Identifies the first e-mail address for the contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R125");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R125
            Site.CaptureRequirementIfAreEqual<EmailAddressKeyType>(
                EmailAddressKeyType.EmailAddress2,
                contacts[1].EmailAddresses[0].Key,
                125,
                @"[In t:EmailAddressKeyType Simple Type] EmailAddress2: Identifies the second e-mail address for the contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R126");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R126
            Site.CaptureRequirementIfAreEqual<EmailAddressKeyType>(
                EmailAddressKeyType.EmailAddress3,
                contacts[2].EmailAddresses[0].Key,
                126,
                @"[In t:EmailAddressKeyType Simple Type] EmailAddress3: Identifies the third e-mail address for the contact (2).");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the FileAsMappingType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC04_VerifyContactItemWithFileAsMappingTypeEnums()
        {
            // The value count of enumeration "FileAsMappingType" is 13.
            int enumCount = 13;

            FileAsMappingType[] fileAsMappingTypes = new FileAsMappingType[enumCount];
            ContactItemType[] requestItems = new ContactItemType[enumCount];
            fileAsMappingTypes[0] = FileAsMappingType.None;
            fileAsMappingTypes[1] = FileAsMappingType.LastCommaFirst;
            fileAsMappingTypes[2] = FileAsMappingType.FirstSpaceLast;
            fileAsMappingTypes[3] = FileAsMappingType.Company;
            fileAsMappingTypes[4] = FileAsMappingType.LastCommaFirstCompany;
            fileAsMappingTypes[5] = FileAsMappingType.CompanyLastFirst;
            fileAsMappingTypes[6] = FileAsMappingType.LastFirst;
            fileAsMappingTypes[7] = FileAsMappingType.LastFirstCompany;
            fileAsMappingTypes[8] = FileAsMappingType.CompanyLastCommaFirst;
            fileAsMappingTypes[9] = FileAsMappingType.LastFirstSuffix;
            fileAsMappingTypes[10] = FileAsMappingType.LastSpaceFirstCompany;
            fileAsMappingTypes[11] = FileAsMappingType.CompanyLastSpaceFirst;
            fileAsMappingTypes[12] = FileAsMappingType.LastSpaceFirst;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain a FileAsMappingType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                FileAsMappingType fileAsMappingType = fileAsMappingTypes[i];

                #region Step 1:Create the contact item with FileAsMappingType
                // Create a contact item with FileAsMappingType.
                requestItems[i] = this.CreateFileAsMappingTypeContact(fileAsMappingType);

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(requestItems[i]);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item.
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsTrue(
                    contacts[i].FileAsMappingSpecified,
                    "The FileAsMapping element in returned contact item should be specified.");

                Site.Assert.IsNotNull(
                    contacts[i].CompleteName,
                    "The returned CompleteName element in contact item should not be null.");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R130, Expected result: FileAsMapping is {0}, FileAs is {1}. Actual result: FileAsMapping is {2}, FileAs is {3}", requestItems[0].FileAsMapping, requestItems[0].FileAs, contacts[0].FileAsMapping, contacts[0].FileAs);

            // If the FileAsMapping element of contacts[0] equals to None, 
            // and the FileAs element in response equals that in request, which is not constructed by other properties,
            // then this requirement can be verified.
            bool isVerifyR130 = contacts[0].FileAsMapping == FileAsMappingType.None
                && contacts[0].FileAs == requestItems[0].FileAs;

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R130
            Site.CaptureRequirementIfIsTrue(
                isVerifyR130,
                130,
                @"[In t:FileAsMappingType Simple Type] None: Indicates that the FileAs value is not constructed from properties of other contacts (2), but is represented by a string, saved ""as is"".  ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R133, Expected result: FileAsMapping is {0}, FileAs is {1}. Actual result: FileAsMapping is {2}, FileAs is {3}", requestItems[3].FileAsMapping, requestItems[3].CompanyName, contacts[3].FileAsMapping, contacts[3].FileAs);

            // If the FileAsMapping element of contacts[3] equals to Company, 
            // and the FileAs element equals to the company name,
            // then this requirement can be verified.
            bool isVerifyR133 = contacts[3].FileAsMapping == FileAsMappingType.Company
                && contacts[3].FileAs == contacts[3].CompanyName;

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R133
            Site.CaptureRequirementIfIsTrue(
                isVerifyR133,
                133,
                @"[In t:FileAsMappingType Simple Type] Company: Indicates that the company name is displayed.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the ImAddressKeyType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC05_VerifyContactItemWithImAddressKeyTypeEnums()
        {
            // The value count of enumeration "ImAddressKeyType" is 3.
            int enumCount = 3;
            ImAddressKeyType[] addressKeyTypes = new ImAddressKeyType[enumCount];

            addressKeyTypes[0] = ImAddressKeyType.ImAddress1;
            addressKeyTypes[1] = ImAddressKeyType.ImAddress2;
            addressKeyTypes[2] = ImAddressKeyType.ImAddress3;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain an ImAddressKeyType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                ImAddressKeyType addressKeyType = addressKeyTypes[i];

                #region Step 1:Create the contact item with ImAddressKeyType
                // Create a contact item with ImAddressKeyType.
                ContactItemType item = this.BuildContactItemWithImAddress(addressKeyType);
                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item.
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsNotNull(
                    contacts[i].ImAddresses,
                    "The ImAddresses element in returned contact item should not be null.");

                Site.Assert.AreEqual<int>(
                    1,
                    contacts[i].ImAddresses.GetLength(0),
                    "One entry of ImAddresses element should be returned!");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R151");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R151
            Site.CaptureRequirementIfAreEqual<ImAddressKeyType>(
                ImAddressKeyType.ImAddress1,
                contacts[0].ImAddresses[0].Key,
                151,
                @"[In t:ImAddressKeyType Simple Type] ImAddress1: Identifies the first instant messaging address for the user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R152");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R152
            Site.CaptureRequirementIfAreEqual<ImAddressKeyType>(
                ImAddressKeyType.ImAddress2,
                contacts[1].ImAddresses[0].Key,
                152,
                @"[In t:ImAddressKeyType Simple Type] ImAddress2: Identifies the second instant messaging address for the user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R153");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R153
            Site.CaptureRequirementIfAreEqual<ImAddressKeyType>(
                ImAddressKeyType.ImAddress3,
                contacts[2].ImAddresses[0].Key,
                153,
                @"[In t:ImAddressKeyType Simple Type] ImAddress3: Identifies the third instant messaging address for the user.");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the PhoneNumberKeyType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC06_VerifyContactItemWithPhoneNumbersKeyTypeEnums()
        {
            // The value count of enumeration "PhoneNumberKeyType" is 19.
            int enumCount = 19;
            PhoneNumberKeyType[] phoneNumberKeyTypes = new PhoneNumberKeyType[enumCount];

            phoneNumberKeyTypes[0] = PhoneNumberKeyType.AssistantPhone;
            phoneNumberKeyTypes[1] = PhoneNumberKeyType.BusinessFax;
            phoneNumberKeyTypes[2] = PhoneNumberKeyType.BusinessPhone;
            phoneNumberKeyTypes[3] = PhoneNumberKeyType.BusinessPhone2;
            phoneNumberKeyTypes[4] = PhoneNumberKeyType.Callback;
            phoneNumberKeyTypes[5] = PhoneNumberKeyType.CarPhone;
            phoneNumberKeyTypes[6] = PhoneNumberKeyType.CompanyMainPhone;
            phoneNumberKeyTypes[7] = PhoneNumberKeyType.HomeFax;
            phoneNumberKeyTypes[8] = PhoneNumberKeyType.HomePhone;
            phoneNumberKeyTypes[9] = PhoneNumberKeyType.HomePhone2;
            phoneNumberKeyTypes[10] = PhoneNumberKeyType.Isdn;
            phoneNumberKeyTypes[11] = PhoneNumberKeyType.MobilePhone;
            phoneNumberKeyTypes[12] = PhoneNumberKeyType.OtherFax;
            phoneNumberKeyTypes[13] = PhoneNumberKeyType.OtherTelephone;
            phoneNumberKeyTypes[14] = PhoneNumberKeyType.Pager;
            phoneNumberKeyTypes[15] = PhoneNumberKeyType.PrimaryPhone;
            phoneNumberKeyTypes[16] = PhoneNumberKeyType.RadioPhone;
            phoneNumberKeyTypes[17] = PhoneNumberKeyType.Telex;
            phoneNumberKeyTypes[18] = PhoneNumberKeyType.TtyTddPhone;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain a PhoneNumberKeyType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                PhoneNumberKeyType phoneNumberKeyType = phoneNumberKeyTypes[i];

                // Create a contact item.
                ContactItemType item = new ContactItemType()
                {
                    // Set a single phone number for the contact.
                    PhoneNumbers = new PhoneNumberDictionaryEntryType[]
                    {
                        new PhoneNumberDictionaryEntryType()
                        {
                            Value = "123456789",
                            Key = phoneNumberKeyType,
                        }
                    }
                };

                #region Step 1:Create the contact item with PhoneNumberKeyType
                // Call CreateItem operation.
                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsNotNull(
                    contacts[i].PhoneNumbers,
                    "The PhoneNumbers element in returned contact item should not be null.");

                Site.Assert.AreEqual<int>(
                    1,
                    contacts[i].PhoneNumbers.GetLength(0),
                    "One entry of PhoneNumbers element should be returned!");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R157");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R157
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.AssistantPhone,
                contacts[0].PhoneNumbers[0].Key,
                157,
                @"[In t:PhoneNumberKeyType Simple Type] AssistantPhone: Identifies the telephone number as the assistant's telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R158");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R158
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.BusinessFax,
                contacts[1].PhoneNumbers[0].Key,
                158,
                @"[In t:PhoneNumberKeyType Simple Type] BusinessFax: Identifies the telephone number as a business fax number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R159");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R159
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.BusinessPhone,
                contacts[2].PhoneNumbers[0].Key,
                159,
                @"[In t:PhoneNumberKeyType Simple Type] BusinessPhone: Identifies the telephone number as a business telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R160");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R160
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.BusinessPhone2,
                contacts[3].PhoneNumbers[0].Key,
                160,
                @"[In t:PhoneNumberKeyType Simple Type] BusinessPhone2: Identifies the telephone number as a second business telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R161");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R161
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.Callback,
                contacts[4].PhoneNumbers[0].Key,
                161,
                @"[In t:PhoneNumberKeyType Simple Type] Callback: Identifies the telephone number as a callback number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R162");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R162
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.CarPhone,
                contacts[5].PhoneNumbers[0].Key,
                162,
                @"[In t:PhoneNumberKeyType Simple Type] CarPhone: Identifies the telephone number as a car telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R163");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R163
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.CompanyMainPhone,
                contacts[6].PhoneNumbers[0].Key,
                163,
                @"[In t:PhoneNumberKeyType Simple Type] CompanyMainPhone: Identifies the telephone number as the company's main telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R164");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R164
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.HomeFax,
                contacts[7].PhoneNumbers[0].Key,
                164,
                @"[In t:PhoneNumberKeyType Simple Type] HomeFax: Identifies the telephone number as a home fax number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R165");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R165
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.HomePhone,
                contacts[8].PhoneNumbers[0].Key,
                165,
                @"[In t:PhoneNumberKeyType Simple Type] HomePhone: Identifies the telephone number as a home telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R166");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R166
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.HomePhone2,
                contacts[9].PhoneNumbers[0].Key,
                166,
                @"[In t:PhoneNumberKeyType Simple Type] HomePhone2: Identifies the telephone number as a second home telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R167");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R167
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.Isdn,
                contacts[10].PhoneNumbers[0].Key,
                167,
                @"[In t:PhoneNumberKeyType Simple Type] Isdn: Identifies the telephone number as an Integrated Services Digital Network (ISDN) line.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R168");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R168
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.MobilePhone,
                contacts[11].PhoneNumbers[0].Key,
                168,
                @"[In t:PhoneNumberKeyType Simple Type] MobilePhone: Identifies the telephone number as a mobile phone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R169");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R169
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.OtherFax,
                contacts[12].PhoneNumbers[0].Key,
                169,
                @"[In t:PhoneNumberKeyType Simple Type] OtherFax: Identifies the telephone number as another fax number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R170");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R170
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.OtherTelephone,
                contacts[13].PhoneNumbers[0].Key,
                170,
                @"[In t:PhoneNumberKeyType Simple Type] OtherTelephone: Identifies the telephone number as another telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R171");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R171
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.Pager,
                contacts[14].PhoneNumbers[0].Key,
                171,
                @"[In t:PhoneNumberKeyType Simple Type] Pager: Identifies the telephone number as a pager.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R172");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R172
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.PrimaryPhone,
                contacts[15].PhoneNumbers[0].Key,
                172,
                @"[In t:PhoneNumberKeyType Simple Type] PrimaryPhone: Identifies the telephone number as the primary telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R173");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R173
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.RadioPhone,
                contacts[16].PhoneNumbers[0].Key,
                173,
                @"[In t:PhoneNumberKeyType Simple Type] RadioPhone: Identifies the telephone number as a radio telephone.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R174");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R174
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.Telex,
                contacts[17].PhoneNumbers[0].Key,
                174,
                @"[In t:PhoneNumberKeyType Simple Type] Telex: Identifies the telephone number as a telex telephone number.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R175");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R175
            Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                PhoneNumberKeyType.TtyTddPhone,
                contacts[18].PhoneNumbers[0].Key,
                175,
                @"[In t:PhoneNumberKeyType Simple Type] TtyTddPhone: Identifies the telephone number as a teletype/telecommunication device for the deaf (TTY/TDD) telephone number.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the PhysicalAddressIndexType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC07_VerifyContactItemWithPhysicalAddressIndexTypeEnums()
        {
            // The value count of enumeration "PhysicalAddressIndexType" is 4.
            int enumCount = 4;
            PhysicalAddressIndexType[] physicalAddressIndexTypes = new PhysicalAddressIndexType[enumCount];

            physicalAddressIndexTypes[0] = PhysicalAddressIndexType.None;
            physicalAddressIndexTypes[1] = PhysicalAddressIndexType.Business;
            physicalAddressIndexTypes[2] = PhysicalAddressIndexType.Home;
            physicalAddressIndexTypes[3] = PhysicalAddressIndexType.Other;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain a PhysicalAddressIndexType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                PhysicalAddressIndexType physicalAddressIndexType = physicalAddressIndexTypes[i];

                #region Step 1:Create the contact item with PhysicalAddressIndexType
                // Create a contact item with PhysicalAddressIndexType.
                ContactItemType item = new ContactItemType()
                {
                    PostalAddressIndex = physicalAddressIndexType,
                    PostalAddressIndexSpecified = true,
                };
                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item.
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsTrue(
                    contacts[i].PostalAddressIndexSpecified,
                    "The PostalAddressIndex element in returned contact item should be present.");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R180");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R180
            Site.CaptureRequirementIfAreEqual<PhysicalAddressIndexType>(
                PhysicalAddressIndexType.None,
                contacts[0].PostalAddressIndex,
                180,
                @"[In t:PhysicalAddressIndexType Simple Type] None: Indicates that no type is specified for the address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R181");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R181
            Site.CaptureRequirementIfAreEqual<PhysicalAddressIndexType>(
                PhysicalAddressIndexType.Business,
                contacts[1].PostalAddressIndex,
                181,
                @"[In t:PhysicalAddressIndexType Simple Type] Business: Displays the address as a business address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R182");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R182
            Site.CaptureRequirementIfAreEqual<PhysicalAddressIndexType>(
                PhysicalAddressIndexType.Home,
                contacts[2].PostalAddressIndex,
                182,
                @"[In t:PhysicalAddressIndexType Simple Type] Home: Displays the address as a home address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R183");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R183
            Site.CaptureRequirementIfAreEqual<PhysicalAddressIndexType>(
                PhysicalAddressIndexType.Other,
                contacts[3].PostalAddressIndex,
                183,
                @"[In t:PhysicalAddressIndexType Simple Type] Other: Displays the address as an address of type other.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the PhysicalAddressKeyType enumeration for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC08_VerifyContactItemWithPhysicalAddressKeyTypeEnums()
        {
            // The value count of enumeration "PhysicalAddressKeyType" is 3.
            int enumCount = 3;
            PhysicalAddressKeyType[] physicalAddressKeyTypes = new PhysicalAddressKeyType[enumCount];

            physicalAddressKeyTypes[0] = PhysicalAddressKeyType.Business;
            physicalAddressKeyTypes[1] = PhysicalAddressKeyType.Home;
            physicalAddressKeyTypes[2] = PhysicalAddressKeyType.Other;

            // Define a contact array to store the contact items got from GetItem operation response.
            // Each contact should contain a PhysicalAddressKeyType value as its element's value.
            ContactItemType[] contacts = new ContactItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                PhysicalAddressKeyType physicalAddressKeyType = physicalAddressKeyTypes[i];

                #region Step 1:Create the contact item with PhysicalAddressKeyType
                // Create a contact item with PhysicalAddressKeyType.
                ContactItemType item = this.BuildContactItemWithPhysicalAddress(physicalAddressKeyType);
                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

                // Check the response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
                #endregion

                #region Step 2:Get the contact item.
                // The contact item to get.
                ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                GetItemResponseType getItemResponse = this.CallGetItemOperation(itemIds);

                // Check the response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

                Site.Assert.AreEqual<int>(
                    1,
                    getItems.GetLength(0),
                    "One contact item should be returned!");

                contacts[i] = getItems[0];

                Site.Assert.IsNotNull(
                    contacts[i],
                    "The returned contact item should not be null.");

                Site.Assert.IsNotNull(
                    contacts[i].PhysicalAddresses,
                    "The PhysicalAddresses element in returned contact item should not be null.");

                Site.Assert.AreEqual<int>(
                    1,
                    contacts[i].PhysicalAddresses.GetLength(0),
                    "One entry of PhysicalAddresses element should be returned!");
                #endregion
            }

            #region Capture Code

            this.Site.Assert.IsTrue(
                this.IsSchemaValidated,
                "The schema should be validated! Expected result: {0}, Actual result: {1}",
                true.ToString(),
                this.IsSchemaValidated.ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R187");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R187
            Site.CaptureRequirementIfAreEqual<PhysicalAddressKeyType>(
                PhysicalAddressKeyType.Business,
                contacts[0].PhysicalAddresses[0].Key,
                187,
                @"[In t:PhysicalAddressKeyType Simple Type] Business: Identifies the address as a business address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R188");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R188
            Site.CaptureRequirementIfAreEqual<PhysicalAddressKeyType>(
                PhysicalAddressKeyType.Home,
                contacts[1].PhysicalAddresses[0].Key,
                188,
                @"[In t:PhysicalAddressKeyType Simple Type] Home: Identifies the address as a home address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R189");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R189
            Site.CaptureRequirementIfAreEqual<PhysicalAddressKeyType>(
                PhysicalAddressKeyType.Other,
                contacts[2].PhysicalAddresses[0].Key,
                189,
                @"[In t:PhysicalAddressKeyType Simple Type] Other: Identifies the address as an address of type other.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorCannotCreateContactInNonContactFolder response code for contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC09_VerifyErrorCannotCreateContactInNonContactFolder()
        {
            #region Step 1:Create the contact item.
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the contact item
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ContactItemType[1];

            // Create a contact item without optional elements.
            createItemRequest.Items.Items[0] = this.BuildContactItemWithRequiredProperties();

            // Configure create item to draft folder to trigger the error.
            DistinguishedFolderIdType errorDistinguishedFolderId = new DistinguishedFolderIdType();
            errorDistinguishedFolderId.Id = DistinguishedFolderIdNameType.drafts;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = errorDistinguishedFolderId;
            #endregion

            CreateItemResponseType createItemResponse = this.CONTAdapter.CreateItem(createItemRequest);

            Site.Assert.AreEqual<int>(
                 1,
                 createItemResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createItemResponse.ResponseMessages.Items.GetLength(0));
            #endregion

            #region Capture Code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R262");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R262
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCannotCreateContactInNonContactFolder,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "MS-OXWSCDATA",
                262,
                @"[In m:ResponseCodeType Simple Type]The value ""ErrorCannotCreateContactInNonContactFolder"" specifies that an attempt was made to create a contact in a folder other than the Contacts folder.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem request which contains HasPicture element in contact item.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC10_VerifyHasPictureIsReadonly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1275002, this.Site), "Implementation does not support the HasPicture element.");

            #region Step 1:Create the contact item.
            // Call CreateItem operation.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();
            item.HasPicture = false;
            item.HasPictureSpecified = true;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(item);

            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                81001,
                @"[In t:ContactItemType Complex Type] HasPicture element: This element is read-only for the client.<6>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275002");

            // If step above all pass, R1275002 will be verified.
            this.Site.CaptureRequirement(
                1275002,
                @"[In Appendix C: Product Behavior] Implementation does support element HasPicture. (Exchange 2010 and above follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by create contact item in folder that is not a Contacts folder.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC11_VerifyErrorCannotCreateContactInNonContactFolder()
        {
            #region Step 1:Create the contact item.
            // Call CreateItem operation.
            ContactItemType item = this.BuildContactItemWithRequiredProperties();

            CreateItemType createItemRequest = new CreateItemType();

            #region Config the contact item
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ContactItemType[1];

            // Create a contact item without optional elements.
            createItemRequest.Items.Items[0] = item;

            // Configure the SavedItemFolderId of CreateItem request to specify that the created item is saved under which folder.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType()
            {
                Item = new DistinguishedFolderIdType()
                {
                    Id = DistinguishedFolderIdNameType.inbox,
                }
            };
            #endregion

            // Call CreateItem operation.
            CreateItemResponseType createItemResponse = this.CONTAdapter.CreateItem(createItemRequest);

            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorCannotCreateContactInNonContactFolder,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                301001,
                @"[In CreateItem] tns:CreateItemSoapIn: The contact (2) item MUST be created in a Contacts folder, or ErrorCannotCreateContactInNonContactFolder ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, GetItem and DeleteItem operations for person item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCONT"), TestMethod()]
        public void MSOXWSCONT_S01_TC12_VerifyAbchPersonItemType()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(336002, this.Site), "Implementation does support the AbchPersonItemType complex type.");

            #region Step 1:Create an item
            AbchPersonItemType abchPersonItem = new AbchPersonItemType();

            abchPersonItem.AntiLinkInfo = Common.GenerateResourceName(this.Site, "AntiLinkInfo");
            abchPersonItem.ContactCategories = new string[] {
                                    "test category"
                                };
            abchPersonItem.FavoriteOrderSpecified = true;
            abchPersonItem.FavoriteOrder = 1;
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(abchPersonItem);
            #endregion

            #region Find the item
            FindItemType findRequest = new FindItemType();
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.Default;
            DistinguishedFolderIdType disFolderId = new DistinguishedFolderIdType();
            disFolderId.Id = DistinguishedFolderIdNameType.contacts;
            findRequest.ParentFolderIds = new BaseFolderIdType[1];
            findRequest.ParentFolderIds[0] = disFolderId;
            FindItemResponseType resp = this.CONTAdapter.FindItem(findRequest);

            ArrayOfRealItemsType items = ((FindItemResponseMessageType)resp.ResponseMessages.Items[0]).RootFolder.Item as ArrayOfRealItemsType;
            #endregion

            #region Capture code
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16006");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R216006
            Site.CaptureRequirementIfIsNotNull(
                ((AbchPersonItemType)items.Items[0]).AntiLinkInfo,
                16006,
                @"[In t:AbchPersonItemType Complex Type] AntiLinkInfo element: Specifies an ID of a set of people who MUST NOT be linked together automatically.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16012");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R216012
            Site.CaptureRequirementIfIsNotNull(
                ((AbchPersonItemType)items.Items[0]).ContactCategories,
                16012,
                @"[In t:AbchPersonItemType Complex Type] ContactCategories element: Specifies the categories of groups that this person belongs to.");

            Site.Assert.IsTrue(((AbchPersonItemType)items.Items[0]).FavoriteOrderSpecified, "FavoriteOrderSpecified element should be True.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16022");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R16022
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                ((AbchPersonItemType)items.Items[0]).FavoriteOrder,
                16022,
                @"[In t:AbchPersonItemType Complex Type] FavoriteOrder element: Otherwise [If this value is not 0], a non-zero value means this person is a favorite. ");
            #endregion

            #region Step3:Delete the item
            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.ItemIds = createdItemIds;

            // Configure the enumeration value that specifies how an person item is to be deleted.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            DeleteItemResponseType deleteItemResponse = this.CONTAdapter.DeleteItem(deleteItemRequest);


            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 4:Get the deleted person item
            // Call GetItem operation using the deleted item IDs.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            Site.Assert.AreEqual<int>(
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0));

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                getItemResponse.ResponseMessages.Items[0].ResponseClass,
                string.Format(
                    "Get deleted item should be failed! Expected response code: {0}, actual response code: {1}",
                    ResponseCodeType.ErrorItemNotFound,
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion
        }
        #endregion
    }
}