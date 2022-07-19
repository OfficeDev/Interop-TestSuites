namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using System;
    using System.Collections.ObjectModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The bass class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// Define a collection used to store existing contact items.
        /// </summary>
        private Collection<ItemIdType> existContactItems;
        #endregion

        #region Properties
        /// <summary>
        /// Gets existing contact items.
        /// </summary>
        protected Collection<ItemIdType> ExistContactItems
        { 
            get { return this.existContactItems; }
        }

        /// <summary>
        /// Gets Adapter Instances.
        /// </summary>
        protected IMS_OXWSCONTAdapter CONTAdapter { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the schema validation is successful.
        /// </summary>
        protected bool IsSchemaValidated { get; private set; }
        #endregion

        #region Test case initialize and clean up
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.CONTAdapter = Site.GetAdapter<IMS_OXWSCONTAdapter>();
            ExchangeServiceBinding.ServiceResponseEvent += new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);
            this.InitializeCollection();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            ExchangeServiceBinding.ServiceResponseEvent -= new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);
            if (this.ExistContactItems != null && this.ExistContactItems.Count > 0)
            {
                // Get ItemIds.
                DeleteItemType deleteItemRequest = new DeleteItemType();
                ItemIdType[] itemArray = new ItemIdType[this.ExistContactItems.Count];
                this.ExistContactItems.CopyTo(itemArray, 0);
                deleteItemRequest.ItemIds = itemArray;

                // Configure an enumeration value that describes the item is to be deleted permanently.
                deleteItemRequest.DeleteType = DisposalType.HardDelete;

                DeleteItemResponseType deleteItemResponse = this.CONTAdapter.DeleteItem(deleteItemRequest);

                foreach (ResponseMessageType messageType in deleteItemResponse.ResponseMessages.Items)
                {
                    Site.Assert.AreEqual<ResponseCodeType>(
                        ResponseCodeType.NoError,
                        messageType.ResponseCode,
                        string.Format(
                            "Delete contact item should not fail! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            messageType.ResponseCode));
                }
            }

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();

            base.TestCleanup();
        }
        #endregion

        #region Test case base methods
        /// <summary>
        /// Initialize the collection of existing contact items.
        /// </summary>
        protected void InitializeCollection()
        {
            if (this.existContactItems == null)
            {
                this.existContactItems = new Collection<ItemIdType>();
            }
            else
            {
                this.existContactItems.Clear();
            }
        }

        /// <summary>
        /// Call GetItem operation.
        /// </summary>
        /// <param name="itemIds">The ItemIds to be gotten.</param>
        /// <returns>The GetItem response.</returns>
        protected GetItemResponseType CallGetItemOperation(ItemIdType[] itemIds)
        {
            GetItemType getItem = new GetItemType();
            GetItemResponseType getItemResponse = new GetItemResponseType();

            // Set the ItemShape property of GetItem operation's request
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // The contact item to get.
            getItem.ItemIds = itemIds;

            getItemResponse = this.CONTAdapter.GetItem(getItem);

            return getItemResponse;
        }

        /// <summary>
        /// Call CreateItem operation.
        /// </summary>
        /// <param name="contactItem">The contact item to be created.</param>
        /// <returns>The CreateItem response.</returns>
        protected CreateItemResponseType CallCreateItemOperation(ContactItemType contactItem)
        {
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the contact item
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ContactItemType[1];

            // Create a contact item without optional elements.
            createItemRequest.Items.Items[0] = contactItem;

            // Configure the SavedItemFolderId of CreateItem request to specify that the created item is saved under which folder.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType()
            {
                Item = new DistinguishedFolderIdType()
                {
                    Id = DistinguishedFolderIdNameType.contacts,
                }
            };
            #endregion

            // Call CreateItem operation.
            CreateItemResponseType createItemResponse = this.CONTAdapter.CreateItem(createItemRequest);

            return createItemResponse;
        }

        /// <summary>
        /// Create a contact item with all properties.
        /// </summary>
        /// <returns>The contact item object.</returns>
        protected ContactItemType CreateFullPropertiesContact()
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType()
            {
                // Set the relevant properties on the contact.
                FileAs = Common.GenerateResourceName(
                this.Site, "FirstContact"),
                FileAsMapping = FileAsMappingType.None,
                FileAsMappingSpecified = true,

                // Set a single e-mail address for the contact.
                EmailAddresses = new EmailAddressDictionaryEntryType[]
                {
                    new EmailAddressDictionaryEntryType()
                    {
                        Value = Common.GenerateResourceName(this.Site, "EmailAddress") + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site),
                        Key = EmailAddressKeyType.EmailAddress1,
                    }
                },

                // Set the information of company for the contact.
                Companies = new string[]
                {
                    Common.GenerateResourceName(
                        this.Site,
                        "Company")
                },
                CompanyName = Common.GenerateResourceName(
                        this.Site,
                        "Company"),

                // Set the information of complete name for the contact.
                DisplayName = Common.GenerateResourceName(
                        this.Site,
                        "DisplayName"),
                Nickname = Common.GenerateResourceName(
                        this.Site,
                        "NickName"),
                GivenName = Common.GenerateResourceName(
                        this.Site,
                        "GivenName"),
                MiddleName = Common.GenerateResourceName(
                        this.Site,
                        "MiddleName"),
                Surname = Common.GenerateResourceName(
                        this.Site,
                        "SurName"),
                Initials = Common.GenerateResourceName(
                        this.Site,
                        "CI"),
                Generation = Common.GenerateResourceName(
                        this.Site,
                        "SR."),

                ExtendedProperty = new ExtendedPropertyType[]
                {
                    // Set the title of the complete name.
                    new ExtendedPropertyType()
                    {
                        ExtendedFieldURI = new PathToExtendedFieldType()
                        {
                            PropertyType = MapiPropertyTypeType.String,

                            // The property tag. The PropertyTag attribute can be represented as either a hexadecimal value or a short integer.
                            // The hexadecimal value range: 0x8000< hexadecimal value <0xFFFE, it represents the custom range of properties.
                            // If the PropertyTag attribute is used, the DistinguishedPropertySetId, PropertySetId, PropertyName, and PropertyId attributes must not be used.
                            PropertyTag = "0x3a45",
                        },
                        Item = Common.GenerateResourceName(
                            this.Site,
                            "MR."),
                    },

                    // Set the YomiFirstName of the complete name.
                    new ExtendedPropertyType()
                    {
                        ExtendedFieldURI = new PathToExtendedFieldType()
                        {
                            PropertyType = MapiPropertyTypeType.String,
                            DistinguishedPropertySetId = DistinguishedPropertySetType.Address,

                            // The GUID of the PSETID_Address. The value must be 00062004-0000-0000-C000-000000000046.
                            // The detail refer to:http://go.microsoft.com/fwlink/?LinkId=517862 
                            PropertySetId = "00062004-0000-0000-C000-000000000046",
            
                            // The dispatch ID. The value must be 0x0000802C. The detail refer to:http://go.microsoft.com/fwlink/?LinkId=517863 
                            PropertyId = Convert.ToInt32("0x0000802C", 16),
                            PropertyIdSpecified = true,
                        },
                        Item = Common.GenerateResourceName(
                            this.Site,
                            "YomiFirstName"),
                    },

                    // Set the YomiLastName of the complete name.
                    new ExtendedPropertyType()
                    {
                        ExtendedFieldURI = new PathToExtendedFieldType()
                        {
                            PropertyType = MapiPropertyTypeType.String,
                            DistinguishedPropertySetId = DistinguishedPropertySetType.Address,

                            // The GUID of the PSETID_Address. The value must be 00062004-0000-0000-C000-000000000046. 
                            // The detail refer to:http://go.microsoft.com/fwlink/?LinkId=517862 
                            PropertySetId = "00062004-0000-0000-C000-000000000046",

                            // The dispatch ID. The value must be 0x0000802D. The detail refer to:http://go.microsoft.com/fwlink/?LinkId=517863 
                            PropertyId = Convert.ToInt32("0x0000802D", 16),
                            PropertyIdSpecified = true,
                        },
                        Item = Common.GenerateResourceName(
                            this.Site,
                            "YomiLastName"),
                    },
                },

                // Set the instant messaging addresses for a contact.
                ImAddresses = new ImAddressDictionaryEntryType[]
                    {
                        new ImAddressDictionaryEntryType()
                        {
                            Key = new ImAddressKeyType(),
                            Value = Common.GenerateResourceName(
                                    this.Site,
                                    "DetailAddress"),
                        },
                    },

                // Set the contact telephone number.
                PhoneNumbers = new PhoneNumberDictionaryEntryType[]
                    {
                        new PhoneNumberDictionaryEntryType()
                        {
                            Key = PhoneNumberKeyType.BusinessPhone,
                            Value = "123456789",
                        }
                    },

                // Set a single contact physical address.
                PhysicalAddresses = new PhysicalAddressDictionaryEntryType[]
                {
                    new PhysicalAddressDictionaryEntryType()
                    {
                        City = Common.GenerateResourceName(
                            this.Site,
                            "City"),
                        CountryOrRegion = Common.GenerateResourceName(
                            this.Site,
                            "CountryOrRegion"),
                        PostalCode = "123456",
                        State = Common.GenerateResourceName(
                            this.Site,
                            "State"),
                        Street = Common.GenerateResourceName(
                            this.Site,
                            "Street"),
                        Key = PhysicalAddressKeyType.Business,
                    }
                },

                // Set other relevant properties on the contact.
                SpouseName = Common.GenerateResourceName(
                        this.Site,
                        "SpouseName"),
                AssistantName = Common.GenerateResourceName(
                        this.Site,
                        "AssistantName"),
                Birthday = Convert.ToDateTime("1988-10-16 11:59:00"),
                BirthdayLocal = TimeZoneInfo.ConvertTimeToUtc(Convert.ToDateTime("1988-10-16 11:59:00"), TimeZoneInfo.Local),
                BirthdaySpecified = true,
                BirthdayLocalSpecified = true,
                BusinessHomePage = "http://www.microsoft.com",
                JobTitle = Common.GenerateResourceName(
                        this.Site,
                        "JobTitle"),
                Manager = Common.GenerateResourceName(
                        this.Site,
                        "Manager"),
                Mileage = Common.GenerateResourceName(
                        this.Site,
                        "Mileage"),
                OfficeLocation = Common.GenerateResourceName(
                        this.Site,
                        "OfficeLocation"),
                Subject = Common.GenerateResourceName(
                        this.Site,
                        "Subject"),
                Categories = new string[]
                {
                    Common.GenerateResourceName(
                        this.Site,
                        "Categories")
                },
                HasAttachments = false,
                HasAttachmentsSpecified = false,
                HasPicture = false,
                HasPictureSpecified = false,
                Culture = "en-US",
                Department = Common.GenerateResourceName(
                        this.Site,
                        "Department"),
                ContactSource = ContactSourceType.ActiveDirectory,
                ContactSourceSpecified = true,
                Importance = ImportanceChoicesType.Normal,
                ImportanceSpecified = true,
                Profession = Common.GenerateResourceName(
                        this.Site,
                        "Profession"),
                Sensitivity = SensitivityChoicesType.Normal,
                SensitivitySpecified = true,
                Children = new string[]
                { 
                    Common.GenerateResourceName(
                        this.Site,
                        "Children") 
                },
                PostalAddressIndex = PhysicalAddressIndexType.Business,
                PostalAddressIndexSpecified = true,
                WeddingAnniversary = Convert.ToDateTime("2010-10-10 11:59:00"),
                WeddingAnniversarySpecified = true,
                WeddingAnniversaryLocal = TimeZoneInfo.ConvertTimeToUtc(Convert.ToDateTime("2010-10-10 11:59:00"), TimeZoneInfo.Local),
                WeddingAnniversaryLocalSpecified = true
            };

            if (Common.IsRequirementEnabled(1275004, this.Site))
            {
                item.AccountName = Common.GenerateResourceName(this.Site, "AccountName");
            }

            if (Common.IsRequirementEnabled(1275006, this.Site))
            {
                item.IsAutoUpdateDisabledSpecified = true;
            }

            if (Common.IsRequirementEnabled(1275008, this.Site))
            {
                item.Comment = Common.GenerateResourceName(this.Site, "Comment");
            }

            if (Common.IsRequirementEnabled(1275014, this.Site))
            {
                item.Gender = "Female";
            }

            if (Common.IsRequirementEnabled(1275026, this.Site))
            {
                item.SourceId = Common.GenerateResourceName(this.Site, "SourceId");
            }

            if (Common.IsRequirementEnabled(1275032, this.Site))
            {
                item.Urls = new ContactUrlDictionaryEntryType[]
                {
                    new ContactUrlDictionaryEntryType()
                    {
                        Type = ContactUrlKeyType.Business,
                        Name = "This is a url",
                        Address = Common.GenerateResourceName(
                            this.Site,
                            "Address"),
                    }
                };
            }

            if (Common.IsRequirementEnabled(1275034, this.Site))
            {
                item.CidSpecified = true;
                item.Cid = 1;
            }

            if (Common.IsRequirementEnabled(1275040, this.Site))
            {
                item.SkypeId = Common.GenerateResourceName(this.Site, "SkypeId");
            }

            if (Common.IsRequirementEnabled(1275044, this.Site))
            {
                item.YomiNickname = Common.GenerateResourceName(this.Site, "YomiNickname");
            }

            return item;
        }

        /// <summary>
        /// Create a contact item with enumeration value of FileAsMappingType.
        /// </summary>
        /// <param name="fileAsMappingType">The enumeration value of FileAsMappingType.</param>
        /// <returns>The contact item object.</returns>
        protected ContactItemType CreateFileAsMappingTypeContact(FileAsMappingType fileAsMappingType)
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType()
            {
                // Set the relevant properties on the contact.
                FileAsMapping = fileAsMappingType,
                FileAsMappingSpecified = true,

                // Set the information of complete name for the contact.
                DisplayName = Common.GenerateResourceName(
                        this.Site,
                        "DisplayName"),
                GivenName = Common.GenerateResourceName(
                        this.Site,
                        "GivenName"),
                MiddleName = Common.GenerateResourceName(
                        this.Site,
                        "MiddleName"),
                Surname = Common.GenerateResourceName(
                        this.Site,
                        "SurName"),
                Generation = Common.GenerateResourceName(
                        this.Site,
                        "SR."),

                // Set the information of company for the contact.
                Companies = new string[]
                {
                    Common.GenerateResourceName(
                        this.Site,
                        "Company")
                },

                CompanyName = Common.GenerateResourceName(
                        this.Site,
                        "Company"),
            };

            if (fileAsMappingType == FileAsMappingType.None)
            {
                item.FileAs = Common.GenerateResourceName(
                this.Site, "FirstContact");
            }

            return item;
        }

        /// <summary>
        /// Build a contact item with enumeration value of EmailAddressKeyType.
        /// </summary>
        /// <param name="emailAddressKeyType">The enumeration value of EmailAddressKeyType.</param>
        /// <returns>The contact item object.</returns>
        protected ContactItemType BuildContactItemWithEmailAddress(EmailAddressKeyType emailAddressKeyType)
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType()
            {
                // Set a single e-mail address for the contact.
                EmailAddresses = new EmailAddressDictionaryEntryType[]
                {
                    new EmailAddressDictionaryEntryType()
                    {
                        Value = Common.GenerateResourceName(this.Site, "EmailAddress") + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site),
                        Key = emailAddressKeyType,
                        Name = Common.GenerateResourceName(this.Site, "Name"),
                    }
                }
            };
            return item;
        }

        /// <summary>
        /// Build a contact item with enumeration value of ImAddressKeyType.
        /// </summary>
        /// <param name="instantMessagingAddress">The enumeration value of ImAddressKeyType.</param>
        /// <returns>The contact item object.</returns>
        protected ContactItemType BuildContactItemWithImAddress(ImAddressKeyType instantMessagingAddress)
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType()
            {
                // Set a single IM address for the contact.
                ImAddresses = new ImAddressDictionaryEntryType[]
                {
                    new ImAddressDictionaryEntryType()
                    {
                        Value = Common.GenerateResourceName(this.Site, "EmailAddress") + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site),
                        Key = instantMessagingAddress,
                    }
                }
            };

            return item;
        }

        /// <summary>
        /// Build a contact item with enumeration value of PhysicalAddressKeyType.
        /// </summary>
        /// <param name="physicalAddressKeyType">The enumeration value of PhysicalAddressKeyType.</param>
        /// <returns>The contact item object.</returns>
        protected ContactItemType BuildContactItemWithPhysicalAddress(PhysicalAddressKeyType physicalAddressKeyType)
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType()
            {
                // Set a single contact physical address.
                PhysicalAddresses = new PhysicalAddressDictionaryEntryType[]
                {
                    new PhysicalAddressDictionaryEntryType()
                    {
                        City = Common.GenerateResourceName(
                            this.Site,
                            "City"),
                        Key = physicalAddressKeyType,
                    }
                 }
            };

            return item;
        }

        /// <summary>
        /// Build a contact item with only required elements.
        /// </summary>
        /// <returns>The contact item object.</returns>
        protected ContactItemType BuildContactItemWithRequiredProperties()
        {
            // Create a contact item type.
            ContactItemType item = new ContactItemType();
            item.FileAs = Common.GenerateResourceName(
                this.Site, "FirstContact");
            return item;
        }

        /// <summary>
        /// Create item with minimum elements which are needed.
        /// </summary>
        /// <param name="item">The item to be created.</param>
        /// <returns>The ItemId of the created item.</returns>
        protected ItemIdType[] CreateItemWithMinimumElements(AbchPersonItemType item)
        {
            CreateItemType createItemRequest = new CreateItemType();

            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new AbchPersonItemType[1];

            // Create a person item without optional elements.
            createItemRequest.Items.Items[0] = item;

            // Configure the SavedItemFolderId of CreateItem request to specify that the created item is saved under which folder.
            createItemRequest.SavedItemFolderId = new TargetFolderIdType()
            {
                Item = new DistinguishedFolderIdType()
                {
                    Id = DistinguishedFolderIdNameType.contacts,
                }
            };

            CreateItemResponseType createItemResponse = this.CONTAdapter.CreateItem(createItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 createdItemIds.GetLength(0),
                 "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdItemIds.GetLength(0));

            return createdItemIds;
        }
        #endregion

        #region Capture methods
        #region Verify requirements related to child element of ContactItemType complex types
        /// <summary>
        /// Capture child element of ContactItemType Complex Type related requirements.
        /// </summary>
        /// <param name="responseContactItem">A contact item from the response package of GetItem operation.</param>
        /// <param name="requestContactItem">A contact item from the request package of CreateItem operation.</param>
        protected void VerifyChildElementInContactItemTypeComplexType(ContactItemType responseContactItem, ContactItemType requestContactItem)
        {
            Site.Assert.IsNotNull(responseContactItem, "The contact item should not be null!");

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema validation result should be true.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R20");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R20
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.FileAs,
                20,
                @"[In t:ContactItemType Complex Type] The type of the element of FileAs is xs:string [XMLSCHEMA2]");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R24");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R24
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.DisplayName,
                24,
                @"[In t:ContactItemType Complex Type] The type of the element of DisplayName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R25");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R25
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.DisplayName,
                responseContactItem.DisplayName,
                25,
                @"[In t:ContactItemType Complex Type] DisplayName element: Contains the display name of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R26");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R26
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.GivenName,
                26,
                @"[In t:ContactItemType Complex Type] The type of the element of GivenName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R27");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R27
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.GivenName,
                responseContactItem.GivenName,
                27,
                @"[In t:ContactItemType Complex Type] GivenName element: Contains the given name for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R28");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R28
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Initials,
                28,
                @"[In t:ContactItemType Complex Type] The type of the element of Initials is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R29");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R29
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Initials,
                responseContactItem.Initials,
                29,
                @"[In t:ContactItemType Complex Type] Initials element: Contains the initials for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R30");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R30
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.MiddleName,
                30,
                @"[In t:ContactItemType Complex Type] The type of the element of MiddleName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R31");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R31
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.MiddleName,
                responseContactItem.MiddleName,
                31,
                @"[In t:ContactItemType Complex Type] MiddleName element: Represents the middle name of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R32");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R32
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Nickname,
                32,
                @"[In t:ContactItemType Complex Type] The type of the element of Nickname is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R33");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R33
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Nickname,
                responseContactItem.Nickname,
                33,
                @"[In t:ContactItemType Complex Type] Nickname element: Represents the nickname of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R36");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R36
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.CompanyName,
                36,
                @"[In t:ContactItemType Complex Type] The type of the element of CompanyName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R37");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R37
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.CompanyName,
                responseContactItem.CompanyName,
                37,
                @"[In t:ContactItemType Complex Type] CompanyName element: Contains the company name that is associated with a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R44");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R44
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.AssistantName,
                44,
                @"[In t:ContactItemType Complex Type] The type of the element of AssistantName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R45");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R45
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.AssistantName,
                responseContactItem.AssistantName,
                45,
                @"[In t:ContactItemType Complex Type] AssistantName element: Contains the name of the assistant for the contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R46");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R46
            Site.CaptureRequirementIfIsTrue(
                responseContactItem.BirthdaySpecified,
                46,
                @"[In t:ContactItemType Complex Type] The type of the element of Birthday is xs:dateTime [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R47");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R47
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                requestContactItem.Birthday,
                responseContactItem.Birthday,
                47,
                @"[In t:ContactItemType Complex Type] Birthday element: Represents the birthday of the contact (2).");
            if (Common.IsRequirementEnabled(12750102, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R4701");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R4701
                Site.CaptureRequirementIfIsTrue(
                    responseContactItem.BirthdayLocalSpecified,
                    4701,
                    @"[In t:ContactItemType Complex Type] The type of the element of BirthdayLocal is xs:dateTime.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R4702");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R4702
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    requestContactItem.BirthdayLocal,
                    responseContactItem.BirthdayLocal,
                    4702,
                    @"[In t:ContactItemType Complex Type] BirthdayLocal element: Represents the birthday of the contact (2) in the client's local time zone.<6>");

            }
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R56");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R56
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Department,
                56,
                @"[In t:ContactItemType Complex Type] The type of the element of Department is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R57");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R57
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Department,
                responseContactItem.Department,
                57,
                @"[In t:ContactItemType Complex Type] Department element: Contains the work department for the contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R58");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R58
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Generation,
                58,
                @"[In t:ContactItemType Complex Type] The type of the element of Generation is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R59");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R59
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Generation,
                responseContactItem.Generation,
                59,
                @"[In t:ContactItemType Complex Type] Generation element: Contains a generational abbreviation that follows the full name of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R62");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R62
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.JobTitle,
                62,
                @"[In t:ContactItemType Complex Type] The type of the element of JobTitle is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R63");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R63
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.JobTitle,
                responseContactItem.JobTitle,
                63,
                @"[In t:ContactItemType Complex Type] JobTitle element: Contains the job title of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R64");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R64
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Manager,
                64,
                @"[In t:ContactItemType Complex Type] The type of the element of Manager is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R65");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R65
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Manager,
                responseContactItem.Manager,
                65,
                @"[In t:ContactItemType Complex Type] Manager element: Represents the manager of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R66");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R66
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Mileage,
                66,
                @"[In t:ContactItemType Complex Type] The type of the element of Mileage is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R67");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R67
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Mileage,
                responseContactItem.Mileage,
                67,
                @"[In t:ContactItemType Complex Type] Mileage element: Represents the mileage for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R68");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R68
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.OfficeLocation,
                68,
                @"[In t:ContactItemType Complex Type] The type of the element of OfficeLocation is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R69");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R69
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.OfficeLocation,
                responseContactItem.OfficeLocation,
                69,
                @"[In t:ContactItemType Complex Type] OfficeLocation element: Represents the office location of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R72");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R72
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Profession,
                72,
                @"[In t:ContactItemType Complex Type] The type of the element of Profession is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R73");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R73
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Profession,
                responseContactItem.Profession,
                73,
                @"[In t:ContactItemType Complex Type] Profession element: Represents the profession of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R74");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R74
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.SpouseName,
                74,
                @"[In t:ContactItemType Complex Type] The type of the element of SpouseName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R75");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R75
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.SpouseName,
                responseContactItem.SpouseName,
                75,
                @"[In t:ContactItemType Complex Type] SpouseName element: Represents the name of the spouse/partner of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R76");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R76
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Surname,
                76,
                @"[In t:ContactItemType Complex Type] The type of the element of Surname is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R77");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R77
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Surname,
                responseContactItem.Surname,
                77,
                @"[In t:ContactItemType Complex Type] Surname element: Contains the surname of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R78");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R78
            Site.CaptureRequirementIfIsTrue(
                responseContactItem.WeddingAnniversarySpecified,
                78,
                @"[In t:ContactItemType Complex Type] The type of the element of WeddingAnniversary is xs:dateTime");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R79");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R79
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                requestContactItem.WeddingAnniversary,
                responseContactItem.WeddingAnniversary,
                79,
                @"[In t:ContactItemType Complex Type] WeddingAnniversary element: Contains the wedding anniversary date of a contact (2).");
            if (Common.IsRequirementEnabled(12750104, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R7901");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R7901
                Site.CaptureRequirementIfIsTrue(
                    responseContactItem.WeddingAnniversaryLocalSpecified,
                    7901,
                    @"[In t:ContactItemType Complex Type] The type of the element of WeddingAnniversaryLocal is xs:dateTime");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R7902
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    requestContactItem.WeddingAnniversaryLocal,
                    responseContactItem.WeddingAnniversaryLocal,
                    7902,
                    @"[In t:ContactItemType Complex Type] WeddingAnniversaryLocal element: Contains the wedding anniversary date of a contact (2) in the client's local time zone.<7>");

            }
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R22");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R22
            Site.CaptureRequirementIfIsTrue(
                responseContactItem.FileAsMappingSpecified,
                22,
                @"[In t:ContactItemType Complex Type] The type of the element of FileAsMapping is t:FileAsMappingType (section 3.1.4.1.2.4)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R34");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R34
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.CompleteName,
                34,
                @"[In t:ContactItemType Complex Type] The type of the element of CompleteName is t:CompleteNameType (section 3.1.4.1.1.5)");

            // Verify the CompleteNameType element of ContactItemType Complex Type.
            this.VerifyCompleteNameTypeComplexType(responseContactItem, requestContactItem);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R38");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R38
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.EmailAddresses,
                38,
                @"[In t:ContactItemType Complex Type] The type of the element of EmailAddresses is t:EmailAddressDictionaryType (section3.1.4.1.1.11)");

            // Verify the EmailAddressDictionaryType element of ContactItemType Complex Type.
            this.VerifyEmailAddressDictionaryTypeComplexType(responseContactItem.EmailAddresses, requestContactItem.EmailAddresses);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R40");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R40
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.PhysicalAddresses,
                40,
                @"[In t:ContactItemType Complex Type] The type of the element of PhysicalAddresses is t:PhysicalAddressDictionaryType (section 3.1.4.1.1.17).");

            this.VerifylPhysicalAddressDictionaryTypeComplexType(responseContactItem.PhysicalAddresses, requestContactItem.PhysicalAddresses);

            if (Common.IsRequirementEnabled(1275084, this.Site))
            {
                this.VerifyContactUrlDictionaryTypeComplexType(responseContactItem.Urls, requestContactItem.Urls);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R42");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R42
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.PhoneNumbers,
                42,
                @"[In t:ContactItemType Complex Type] The type of the element of PhoneNumbers is t:PhoneNumberDictionaryType (section 3.1.4.1.1.15).");

            // Verify the PhoneNumberDictionaryType element of ContactItemType Complex Type.
            this.VerifyPhoneNumberDictionaryTypeComplexType(responseContactItem.PhoneNumbers, requestContactItem.PhoneNumbers);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R48");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R48
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.BusinessHomePage,
                48,
                @"[In t:ContactItemType Complex Type] The type of the element of BusinessHomePage is xs:anyURI [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R49");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R49
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.BusinessHomePage,
                responseContactItem.BusinessHomePage,
                49,
                @"[In t:ContactItemType Complex Type] BusinessHomePage element: Contains the business home page Uniform Resource Identifier (URI) of a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R60");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R60
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.ImAddresses,
                60,
                @"[In t:ContactItemType Complex Type] The type of the element of ImAddresses is t:ImAddressDictionaryType (section 3.1.4.1.1.13).");

            // Verify the ImAddressDictionaryType element of ContactItemType Complex Type.
            this.VerifyImAddressDictionaryTypeComplexType(responseContactItem.ImAddresses, requestContactItem.ImAddresses);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R70");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R70
            Site.CaptureRequirementIfIsTrue(
                responseContactItem.PostalAddressIndexSpecified,
                70,
                @"[In t:ContactItemType Complex Type] The type of the element of PostalAddressIndex is t:PhysicalAddressIndexType (section 3.1.4.1.2.7)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R71");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R71
            this.Site.CaptureRequirementIfAreEqual<PhysicalAddressIndexType>(
                requestContactItem.PostalAddressIndex,
                responseContactItem.PostalAddressIndex,
                71,
                @"[In t:ContactItemType Complex Type] PostalAddressIndex element: Represents the index of one of the physical addresses, which is a contact's (2) mailing address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R177");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R177
            // Because R177 is captured based on R70 and R61, so if both of them are captured successfully, R177 can be captured.
            this.Site.CaptureRequirement(
                177,
                @"[t:PhysicalAddressIndexType Simple Type] The PhysicalAddressIndexType simple  type identifies the display types for physical addresses.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R50");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R50
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Children,
                50,
                @"[In t:ContactItemType Complex Type] The type of the element of Children is t:ArrayOfStringsType ([MS-OXWSCDATA] section 2.2.4.13).");

            Site.Assert.AreEqual<int>(
                1,
                responseContactItem.Children.Length,
                string.Format(
                    "The children of the item from response should be 1, actual: '{0}'.", responseContactItem.Children.Length));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R51");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R51
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Children[0],
                responseContactItem.Children[0],
                51,
                @"[In t:ContactItemType Complex Type] Children element: Contains the names of children for the contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R52");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R52
            Site.CaptureRequirementIfIsNotNull(
                responseContactItem.Companies,
                52,
                @"[In t:ContactItemType Complex Type] The type of the element of Companies is t:ArrayOfStringsType.");

            Site.Assert.AreEqual<int>(
                1,
                responseContactItem.Companies.Length,
                string.Format(
                    "The companies of the item from response should be 1, actual: '{0}'.", responseContactItem.Companies.Length));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R53");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R53
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Companies[0],
                responseContactItem.Companies[0],
                53,
                @"[In t:ContactItemType Complex Type] Companies element: Contains the names of companies that are associated with a contact (2).");
        }

        #endregion

        #region Verify requirements related to EmailAddressDictionaryEntryType complex types
        /// <summary>
        /// Capture EmailAddressDictionaryEntryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseEmailAddressDictionaryEntryType">The EmailAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestEmailAddressDictionaryEntryType">The EmailAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyEmailAddressDictionaryEntryTypeComplexType(EmailAddressDictionaryEntryType responseEmailAddressDictionaryEntryType, EmailAddressDictionaryEntryType requestEmailAddressDictionaryEntryType)
        {
            Site.Assert.IsNotNull(responseEmailAddressDictionaryEntryType, "The EmailAddresses element should not contain null entry!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R227");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R227
            // Key element is a required element in schema, EmailAddressKeyType is verified when the schema is validated.
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                227,
                @"[In t:EmailAddressDictionaryEntryType Complex Type] The type of the attribute of Key is t:EmailAddressKeyType (section 3.1.4.1.2.3).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R228");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R228
            this.Site.CaptureRequirementIfAreEqual<EmailAddressKeyType>(
                requestEmailAddressDictionaryEntryType.Key,
                responseEmailAddressDictionaryEntryType.Key,
                228,
                @"[In t:EmailAddressDictionaryEntryType Complex Type] Key attribute: Contains a value that identifies an e-mail address that is associated with a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R225");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R225
            // Because R225 is captured based on R227 and R228, so if both of them are captured successfully, R225 also should be captured.
            this.Site.CaptureRequirement(
                225,
                @"[In t:EmailAddressDictionaryEntryType Complex Type] The EmailAddressDictionaryEntryType complex type represents an e-mail address that is associated with a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R238");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R238
            // Because R238 is captured in the same condition with R228, so if R228 is captured successfully, R238 also should be captured.
            this.Site.CaptureRequirement(
                238,
                @"[In t:EmailAddressDictionaryType Complex Type] Entry element: Represents an e-mail address that is associated with a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R39");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R39
            // Because R39 is captured in the same condition with R228, so if R228 is captured successfully, R39 also should be captured.
            this.Site.CaptureRequirement(
                39,
                @"[In t:ContactItemType Complex Type] EmailAddress element: Contains e-mail addresses that are associated with a contact (2).");
        }
        #endregion

        #region Verify requirements related to EmailAddressDictionaryType complex types
        /// <summary>
        /// Capture EmailAddressDictionaryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseEmailAddressDictionaryType">The array of EmailAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestEmailAddressDictionaryType">The array of EmailAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyEmailAddressDictionaryTypeComplexType(EmailAddressDictionaryEntryType[] responseEmailAddressDictionaryType, EmailAddressDictionaryEntryType[] requestEmailAddressDictionaryType)
        {
            Site.Assert.IsNotNull(responseEmailAddressDictionaryType, "The EmailAddresses element should not be null!");

            for (int i = 0; i < responseEmailAddressDictionaryType.Length; i++)
            {
                // Verify the EmailAddressDictionaryEntryType element of EmailAddressDictionaryType Complex Type.
                this.VerifyEmailAddressDictionaryEntryTypeComplexType(responseEmailAddressDictionaryType[i], requestEmailAddressDictionaryType[i]);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R237");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R237
                bool isVerifyR237 = this.IsSchemaValidated && (requestEmailAddressDictionaryType[i] != null);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR237,
                    237,
                    @"[In t:EmailAddressDictionaryType Complex Type] The type of the element of Entry is t:EmailAddressDictionaryEntryType (section 3.1.4.1.1.10)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R235");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R235
            // R235 can be captured directly after each e-mail address entry is validated.
            this.Site.CaptureRequirement(
                235,
                @"[In t:EmailAddressDictionaryType Complex Type] The EmailAddressDictionaryType complex type contains e-mail addresses.");
        }
        #endregion

        #region Verify requirements related to ImAddressDictionaryEntryType complex types
        /// <summary>
        /// Capture ImAddressDictionaryEntryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseInstantMessageAddressDictionaryEntry">The ImAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestInstantMessageAddressDictionaryEntry">The ImAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyImAddressDictionaryEntryTypeComplexType(ImAddressDictionaryEntryType responseInstantMessageAddressDictionaryEntry, ImAddressDictionaryEntryType requestInstantMessageAddressDictionaryEntry)
        {
            Site.Assert.IsNotNull(responseInstantMessageAddressDictionaryEntry, "The ImAddresses element should not contain null entry!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R241");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R241
            // Key element is a required element in schema, ImAddressKeyType is verified when the schema is validated.
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                241,
                @"[In t:ImAddressDictionaryEntryType Complex Type] The type of the attribute of key is t:ImAddressKeyType (3.1.4.1.2.5)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R242");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R242
            this.Site.CaptureRequirementIfAreEqual<ImAddressKeyType>(
                requestInstantMessageAddressDictionaryEntry.Key,
                responseInstantMessageAddressDictionaryEntry.Key,
                242,
                @"[In t:ImAddressDictionaryEntryType Complex Type] t:ImAddressKeyType attribute: Represents the instant messaging addresses for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R239");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R239
            // Because R239 is captured based on R241 and R242, so if both of them are captured successfully, R239 also should be captured.
            this.Site.CaptureRequirement(
                239,
                @"[In t:ImAddressDictionaryEntryType Complex Type] The ImAddressDictionaryEntryType complex type represents a collection of instant messaging addresses for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R61");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R61
            // Because R61 is captured in the same condition with R242, so if R242 is captured successfully, R61 also should be captured.
            this.Site.CaptureRequirement(
                61,
                @"[In t:ContactItemType Complex Type] ImAddresses element: Contains instant messaging addresses for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R246");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R246
            // Because R246 is captured in the same condition with R242, so if R242 is captured successfully, R246 also should be captured.
            this.Site.CaptureRequirement(
                246,
                @"[In t:ImAddressDictionaryType Complex Type] Entry child element: Represents a collection of instant messaging addresses for a contact (2).");
        }
        #endregion

        #region Verify requirements related to ImAddressDictionaryType complex types
        /// <summary>
        /// Capture ImAddressDictionaryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseInstantMessageAddressDictionary">The array of ImAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestInstantMessageAddressDictionary">The array of ImAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyImAddressDictionaryTypeComplexType(ImAddressDictionaryEntryType[] responseInstantMessageAddressDictionary, ImAddressDictionaryEntryType[] requestInstantMessageAddressDictionary)
        {
            Site.Assert.IsNotNull(responseInstantMessageAddressDictionary, "The ImAddresses element should not be null!");

            for (int i = 0; i < responseInstantMessageAddressDictionary.Length; i++)
            {
                // Verify the ImAddressDictionaryEntryType element of ImAddressDictionaryType Complex Type.
                this.VerifyImAddressDictionaryEntryTypeComplexType(responseInstantMessageAddressDictionary[i], requestInstantMessageAddressDictionary[i]);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R245");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R245
                bool isVerifyR245 = this.IsSchemaValidated && (responseInstantMessageAddressDictionary[i] != null);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR245,
                    245,
                    @"[In t:ImAddressDictionaryType Complex Type] The type of the child element of Entry is t:ImAddressDictionaryEntryType (3.1.4.1.1.14)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R243");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R243
            // R243 can be captured directly after each instant messaging address entry is validated.
            this.Site.CaptureRequirement(
                243,
                @"[In t:ImAddressDictionaryType Complex Type] The ImAddressDictionaryType complex type contains instant messaging addresses for a contact (2).");
        }
        #endregion

        #region Verify requirements related to PhoneNumberDictionaryEntryType complex types
        /// <summary>
        /// Capture PhoneNumberDictionaryEntryType Complex Type related requirements.
        /// </summary>
        /// <param name="responsePhoneNumberDictionaryEntryType">The PhoneNumberDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestPhoneNumberDictionaryEntryType">The PhoneNumberDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyPhoneNumberDictionaryEntryTypeComplexType(PhoneNumberDictionaryEntryType responsePhoneNumberDictionaryEntryType, PhoneNumberDictionaryEntryType requestPhoneNumberDictionaryEntryType)
        {
            Site.Assert.IsNotNull(responsePhoneNumberDictionaryEntryType, "The entry in PhoneNumbers element should not be null!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R249");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R249
            // Key element is a required element in schema, PhoneNumberKeyType is verified when the schema is validated.
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                249,
                @"[In t:PhoneNumberDictionaryEntryType Complex Type] The type of the attribute of Key is t:PhoneNumberKeyType (3.1.4.1.2.6).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R250");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R250
            this.Site.CaptureRequirementIfAreEqual<PhoneNumberKeyType>(
                requestPhoneNumberDictionaryEntryType.Key,
                responsePhoneNumberDictionaryEntryType.Key,
                250,
                @"[In t:PhoneNumberDictionaryEntryType Complex Type] Key attribute: Represents types of telephone numbers for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R247");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R247
            // Because R247 is captured based on R249 and R250, so if both of them are captured successfully, R247 also should be captured.
            this.Site.CaptureRequirement(
                247,
                @"[In t:PhoneNumberDictionaryEntryType Complex Type] The PhoneNumberDictionaryEntryType complex type contains a telephone number for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R43");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R43
            // Because R43 is captured in the same condition with R250, so if R250 is captured successfully, R43 also should be captured.
            this.Site.CaptureRequirement(
                43,
                @"[In t:ContactItemType Complex Type] PhoneNumbers element: Represents a collection of telephone numbers for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R254");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R254
            // Because R254 is captured in the same condition with R250, so if R250 is captured successfully, R254 also should be captured.
            this.Site.CaptureRequirement(
                254,
                @"[In t:PhoneNumberDictionaryType Complex Type] Entry child element: Contains a telephone number for a contact (2).");
        }
        #endregion

        #region Verify requirements related to PhoneNumberDictionaryType complex types
        /// <summary>
        /// Capture PhoneNumberDictionaryType Complex Type related requirements.
        /// </summary>
        /// <param name="responsePhoneNumberDictionaryType">The array of PhoneNumberDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestPhoneNumberDictionaryType">The array of PhoneNumberDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyPhoneNumberDictionaryTypeComplexType(PhoneNumberDictionaryEntryType[] responsePhoneNumberDictionaryType, PhoneNumberDictionaryEntryType[] requestPhoneNumberDictionaryType)
        {
            Site.Assert.IsNotNull(responsePhoneNumberDictionaryType, "The PhoneNumbers element should not be null!");

            for (int i = 0; i < responsePhoneNumberDictionaryType.Length; i++)
            {
                // Verify the PhoneNumberDictionaryEntryType element of PhoneNumberDictionaryType Complex Type.
                this.VerifyPhoneNumberDictionaryEntryTypeComplexType(responsePhoneNumberDictionaryType[i], requestPhoneNumberDictionaryType[i]);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R253");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R253
                bool isVerifyR253 = this.IsSchemaValidated && (responsePhoneNumberDictionaryType[i] != null);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR253,
                    253,
                    @"[In t:PhoneNumberDictionaryType Complex Type] The type of the child element of Entry is t:PhoneNumberDictionaryEntryType (3.1.4.1.1.14).");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R251");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R251
            // R251 can be captured directly after each telephone number entry is validated.
            this.Site.CaptureRequirement(
                251,
                @"[In t:PhoneNumberDictionaryType Complex Type] The PhoneNumberDictionaryType complex type represents telephone numbers for a contact (2).");
        }
        #endregion

        #region Verify requirements related to PhysicalAddressDictionaryEntryType complex types
        /// <summary>
        /// Capture PhysicalAddressDictionaryEntryType Complex Type related requirements.
        /// </summary>
        /// <param name="responsePhysicalAddressDictionaryEntryType">The PhysicalAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestPhysicalAddressDictionaryEntryType">The PhysicalAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyPhysicalAddressDictionaryEntryTypeComplexType(PhysicalAddressDictionaryEntryType responsePhysicalAddressDictionaryEntryType, PhysicalAddressDictionaryEntryType requestPhysicalAddressDictionaryEntryType)
        {
            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType, "The PhysicalAddresses element should not contain null entry!");

            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType.Street, "The Street of the PhysicalAddressDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R257");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R257
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                257,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the child element of Street is xs:string [XMLSCHEMA2]");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R258");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R258
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestPhysicalAddressDictionaryEntryType.Street,
                responsePhysicalAddressDictionaryEntryType.Street,
                258,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] Street child element: Contains the street address for a contact (2) item.");

            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType.City, "The City of the PhysicalAddressDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R259");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R259
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                259,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the child element of City is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R260");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R260
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestPhysicalAddressDictionaryEntryType.City,
                responsePhysicalAddressDictionaryEntryType.City,
                260,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] City child element: Contains the city name for a contact (2) item.");

            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType.State, "The State of the PhysicalAddressDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R261");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R261
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                261,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the child element of State is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R262");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R262
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestPhysicalAddressDictionaryEntryType.State,
                responsePhysicalAddressDictionaryEntryType.State,
                262,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] State child element: Contains the state for a contact (2) item.");

            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType.CountryOrRegion, "The CountryOrRegion of the PhysicalAddressDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R263");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R263
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                263,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the child element of CountryOrRegion is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R264");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R264
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestPhysicalAddressDictionaryEntryType.CountryOrRegion,
                responsePhysicalAddressDictionaryEntryType.CountryOrRegion,
                264,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] CountryOrRegion child element: Contains the country or region for a contact (2) item.");

            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryEntryType.PostalCode, "The PostalCode of the PhysicalAddressDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R265");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R265
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                265,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the child element of PostalCode is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R266");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R266
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestPhysicalAddressDictionaryEntryType.PostalCode,
                responsePhysicalAddressDictionaryEntryType.PostalCode,
                266,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] PostalCode child element: Contains the postal code for a contact (2) item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R267");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R267
            // Key element is a required element in schema, PhysicalAddressKeyType is verified when the schema is validated.
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                267,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type of the attribute of Key is t:PhysicalAddressKeyType (section 3.1.4.1.2.8)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R268");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R268
            this.Site.CaptureRequirementIfAreEqual<PhysicalAddressKeyType>(
                requestPhysicalAddressDictionaryEntryType.Key,
                responsePhysicalAddressDictionaryEntryType.Key,
                268,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] Key attribute: Identifies the types of physical addresses for a contact (2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R184");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R184
            // Because R184 is captured based on R267 and R268, so if both of them are captured successfully, R184 can be captured.
            this.Site.CaptureRequirement(
                184,
                @"[In t:PhysicalAddressKeyType Simple Type] The PhysicalAddressKeyType simple type identifies the types of physical addresses for a contact (2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R255");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R255
            // Because R255 is captured based on R258, R260, R262, R264, R266 and R268, so if all of them are captured successfully, R255 also should be captured.
            this.Site.CaptureRequirement(
                255,
                @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The PhysicalAddressDictionaryEntryType complex type contains information that defines a physical address, such as a street address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R272");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R272
            // Because R272 is captured based on R258, R260, R262, R264, R266 and R268, so if all of them are captured successfully, R272 can be captured.
            this.Site.CaptureRequirement(
                272,
                @"[In t:PhysicalAddressDictionaryType Complex Type] entry child element: Contains information that defines a physical address, such as a street address.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R41");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R41
            // Because R41 is captured based on R258, R260, R262, R264, R266 and R268, so if all of them are captured successfully, R41 can be captured.
            this.Site.CaptureRequirement(
                41,
                @"[In t:ContactItemType Complex Type] PhysicalAddresses element: Represents a collection of physical addresses that are associated with a contact.");
        }
        #endregion

        #region Verify requirements related to PhysicalAddressDictionaryType complex types
        /// <summary>
        /// Capture PhysicalAddressDictionaryType Complex Type related requirements.
        /// </summary>
        /// <param name="responsePhysicalAddressDictionaryType">The array of PhysicalAddressDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestPhysicalAddressDictionaryType">The array of PhysicalAddressDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifylPhysicalAddressDictionaryTypeComplexType(PhysicalAddressDictionaryEntryType[] responsePhysicalAddressDictionaryType, PhysicalAddressDictionaryEntryType[] requestPhysicalAddressDictionaryType)
        {
            Site.Assert.IsNotNull(responsePhysicalAddressDictionaryType, "The PhysicalAddresses element should not be null!");

            for (int i = 0; i < responsePhysicalAddressDictionaryType.Length; i++)
            {
                // Verify the PhysicalAddressDictionaryEntryType element of PhysicalAddressDictionaryType Complex Type.
                this.VerifyPhysicalAddressDictionaryEntryTypeComplexType(responsePhysicalAddressDictionaryType[i], requestPhysicalAddressDictionaryType[i]);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R271");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R271
                bool isVerifyR271 = this.IsSchemaValidated && (responsePhysicalAddressDictionaryType[i] != null);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR271,
                    271,
                    @"[In t:PhysicalAddressDictionaryType Complex Type] The type of the child element of entry is t:PhysicalAddressDictionaryEntryType (section 3.1.4.1.1.16).");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R269");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R269
            // R269 can be captured directly after each physical address entry is validated.
            this.Site.CaptureRequirement(
                269,
                @"[In t:PhysicalAddressDictionaryType Complex Type] The PhysicalAddressDictionaryType complex type contains physical addresses that are associated with a contact (2).");
        }
        #endregion

        #region Verify requirements related to CompleteNameType complex types
        /// <summary>
        /// Capture CompleteNameType Complex Type related requirements.
        /// </summary>
        /// <param name="responseContactItem">The ContactItemType element from the response package of GetItem operation.</param>
        /// <param name="requestContactItem">The ContactItemType element from the request package of CreateItem operation.</param>
        private void VerifyCompleteNameTypeComplexType(ContactItemType responseContactItem, ContactItemType requestContactItem)
        {
            Site.Assert.IsNotNull(responseContactItem.CompleteName, "The complete name element should not be null!");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.Title, "The Title of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R193");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R193
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                193,
                @"[In t:CompleteNameType Complex Type] The type of the element of Title is xs:string [XMLSCHEMA2]");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R194");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R194
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.ExtendedProperty[0].Item as string,
                responseContactItem.CompleteName.Title,
                194,
                @"[In t:CompleteNameType Complex Type] Title element: Contains the title of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.FirstName, "The FirstName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R195");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R195
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                195,
                @"[In t:CompleteNameType Complex Type] The type of the element of FirstName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R355");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R355
            this.Site.CaptureRequirementIfAreEqual<string>(
                responseContactItem.GivenName,
                responseContactItem.CompleteName.FirstName,
                355,
                @"[In t:CompleteNameType Complex Type] This [FirstName element] is the same as GivenName.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R196");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R196
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.GivenName,
                responseContactItem.CompleteName.FirstName,
                196,
                @"[In t:CompleteNameType Complex Type] FirstName element: Contains the first name of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.MiddleName, "The MiddleName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R197");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R197
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                197,
                @"[In t:CompleteNameType Complex Type] The type of the element of MiddleName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R198");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R198
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.MiddleName,
                responseContactItem.CompleteName.MiddleName,
                198,
                @"[In t:CompleteNameType Complex Type] MiddleName element: Contains the middle name of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.LastName, "The LastName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R199");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R199
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                199,
                @"[In t:CompleteNameType Complex Type] The type of the element of LastName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R356");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R356
            this.Site.CaptureRequirementIfAreEqual<string>(
                responseContactItem.Surname,
                responseContactItem.CompleteName.LastName,
                356,
                @"[In t:CompleteNameType Complex Type] This [LastName element] is the same as the Surname.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R200");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R200
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Surname,
                responseContactItem.CompleteName.LastName,
                200,
                @"[In t:CompleteNameType Complex Type] LastName element: Contains the last name of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.Suffix, "The Suffix of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R201");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R201
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                201,
                @"[In t:CompleteNameType Complex Type] The type of the element of Suffix is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R357");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R357
            this.Site.CaptureRequirementIfAreEqual<string>(
                responseContactItem.Generation,
                responseContactItem.CompleteName.Suffix,
                357,
                @"[In t:CompleteNameType Complex Type] This [Suffix element] is the same as the Generation property.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R202");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R202
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Generation,
                responseContactItem.CompleteName.Suffix,
                202,
                @"[In t:CompleteNameType Complex Type] Suffix element: Contains a suffix to a contact's (2) name.");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.Initials, "The Initials of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R203");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R203
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                203,
                @"[In t:CompleteNameType Complex Type] The type of the element of Initials is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R204");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R204
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Initials,
                responseContactItem.CompleteName.Initials,
                204,
                @"[In t:CompleteNameType Complex Type] Initials element: Contains the initials of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.FullName, "The FullName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R205");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R205
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                205,
                @"[In t:CompleteNameType Complex Type] The type of the element of FullName is xs:string");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.Nickname, "The Nickname of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R207");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R207
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                207,
                @"[In t:CompleteNameType Complex Type] The type of the element of Nickname is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R208");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R208
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.Nickname,
                responseContactItem.CompleteName.Nickname,
                208,
                @"[In t:CompleteNameType Complex Type] Nickname element: Contains the nickname of a contact (2).");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.YomiFirstName, "The YomiFirstName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R209");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R209
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                209,
            @"[In t:CompleteNameType Complex Type] The type of the element of YomiFirstName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R210");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R210
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.ExtendedProperty[1].Item as string,
                responseContactItem.CompleteName.YomiFirstName,
                210,
                @"[In t:CompleteNameType Complex Type] YomiFirstName element: Contains the name used in Japan for the searchable or phonetic spelling of a Japanese first name.");

            Site.Assert.IsNotNull(responseContactItem.CompleteName.YomiLastName, "The YomiLastName of CompleteName from response contact item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R211");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R211
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                211,
                @"[In t:CompleteNameType Complex Type] The type of the element of YomiLastName is xs:string");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R212");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R212
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactItem.ExtendedProperty[2].Item as string,
                responseContactItem.CompleteName.YomiLastName,
                212,
                @"[In t:CompleteNameType Complex Type] YomiLastName element: Contains the name used in Japan for the searchable or phonetic spelling of a Japanese last name.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R35");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R35
            // Because R35 is captured based on R194, R196, R198, R200, R202, R204, R208, R210 and R212, so if all of them are captured successfully, R35 can be captured.
            this.Site.CaptureRequirement(
                35,
                @"[In t:ContactItemType Complex Type] CompleteName element: Represents the complete name of a contact (2).");
        }
        #endregion

        #region Verify requirements related to ContactUrlDictionaryEntryType complex types
        /// <summary>
        /// Capture ContactUrlDictionaryEntryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseContactUrlDictionaryEntryType">The ContactUrlDictionaryEntryType element from the response package of GetItem operation.</param>
        /// <param name="requestContactUrlDictionaryEntryType">The ContactUrlDictionaryEntryType element from the request package of CreateItem operation.</param>
        private void VerifyContactUrlDictionaryEntryTypeComplexType(ContactUrlDictionaryEntryType responseContactUrlDictionaryEntryType, ContactUrlDictionaryEntryType requestContactUrlDictionaryEntryType)
        {
            Site.Assert.IsNotNull(responseContactUrlDictionaryEntryType, "The PhysicalAddresses element should not contain null entry!");

            Site.Assert.IsNotNull(responseContactUrlDictionaryEntryType.Type, "The Type of the responseContactUrlDictionaryEntryType element from response should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224004");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224004
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                224004,
                @"[In t:ContactUrlDictionaryEntryType Complex Type] The type of the element Type is t:ContactUrlKeyType (section 3.1.4.1.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224005");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224005
            this.Site.CaptureRequirementIfAreEqual<ContactUrlKeyType>(
                requestContactUrlDictionaryEntryType.Type,
                responseContactUrlDictionaryEntryType.Type,
                224005,
                @"[In t:ContactUrlDictionaryEntryType Complex Type] Type element: specifies the Url type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224006");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224006
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                224006,
                @"[[In t:ContactUrlDictionaryEntryType Complex Type] The type of the element Name is xs:string ([XMLSCHEMA2]).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224007");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224007
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactUrlDictionaryEntryType.Name,
                responseContactUrlDictionaryEntryType.Name,
                224007,
                @"[In t:ContactUrlDictionaryEntryType Complex Type] Name element: Specifies what the url is used for.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224008");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224008
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                224008,
                @"[In t:ContactUrlDictionaryEntryType Complex Type]  The type of the element Address xs:string.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224009");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224009
            this.Site.CaptureRequirementIfAreEqual<string>(
                requestContactUrlDictionaryEntryType.Address,
                responseContactUrlDictionaryEntryType.Address,
                224009,
                @"[In t:ContactUrlDictionaryEntryType Complex Type] Address element: Specifies the Url.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224014");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224014
            // Because R224014 is captured based on R224005 and R224009, so if all of them are captured successfully, R224014 can be captured.
            this.Site.CaptureRequirement(
                224014,
                @"[In t:ContactUrlDictionaryType Complex Type] Url element: Specifies a Url associated with a contact.");
        }
        #endregion

        #region Verify requirements related to ContactUrlDictionaryType complex types
        /// <summary>
        /// Capture ContactUrlDictionaryType Complex Type related requirements.
        /// </summary>
        /// <param name="responseContactUrlDictionaryType">The array of ContactUrlDictionaryType element from the response package of GetItem operation.</param>
        /// <param name="requestContactUrlDictionaryType">The array of ContactUrlDictionaryType element from the request package of CreateItem operation.</param>
        private void VerifyContactUrlDictionaryTypeComplexType(ContactUrlDictionaryEntryType[] responseContactUrlDictionaryType, ContactUrlDictionaryEntryType[] requestContactUrlDictionaryType)
        {
            Site.Assert.IsNotNull(responseContactUrlDictionaryType, "The ContactUrl element should not be null!");

            for (int i = 0; i < responseContactUrlDictionaryType.Length; i++)
            {
                // Verify the PhysicalAddressDictionaryEntryType element of PhysicalAddressDictionaryType Complex Type.
                this.VerifyContactUrlDictionaryEntryTypeComplexType(responseContactUrlDictionaryType[i], requestContactUrlDictionaryType[i]);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224013");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R271
                bool isVerifyR224013 = this.IsSchemaValidated && (responseContactUrlDictionaryType[i] != null);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR224013,
                    224013,
                    @"[In t:ContactUrlDictionaryType Complex Type] The type of the element Url is t:ContactUrlDictionaryEntryType (section 3.1.4.1.1.8).");
            }
        }
        #endregion

        #endregion

        #region Private methods
        /// <summary>
        /// Handle the server response.
        /// </summary>
        /// <param name="request">The request messages.</param>
        /// <param name="response">The response messages.</param>
        /// <param name="isSchemaValidated">Verify the schema.</param>
        private void ExchangeServiceBinding_ResponseEvent(
            BaseRequestType request,
            BaseResponseMessageType response,
            bool isSchemaValidated)
        {
            this.IsSchemaValidated = isSchemaValidated;

            // A flag represents the response contains the item information or not.
            bool hasItemInfo = false;
            if (response is SetUserPhotoResponseMessageType)
            {
                return;
            }

            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                if (responseMessage is ItemInfoResponseMessageType)
                {
                    hasItemInfo = true;
                    break;
                }
            }

            BaseItemIdType[] itemIds;
            if (hasItemInfo)
            {
                itemIds = Common.GetItemIdsFromInfoResponse(response);
            }
            else
            {
                itemIds = new BaseItemIdType[0];
                return;
            }

            foreach (ItemIdType itemId in itemIds)
            {
                bool notExist = true;
                foreach (ItemIdType exist in this.ExistContactItems)
                {
                    if (exist.Id == itemId.Id)
                    {
                        // Check if the Id of the item in response exists in previous recorded item list.
                        notExist = false;
                        break;
                    }
                }

                // If the Id of Item is not exist, add the item into item list.
                if (notExist)
                {
                    this.ExistContactItems.Add(itemId);
                }
            }
        }

        #endregion
    }
}