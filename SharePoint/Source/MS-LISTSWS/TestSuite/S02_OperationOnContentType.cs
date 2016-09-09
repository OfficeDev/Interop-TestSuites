namespace Microsoft.Protocols.TestSuites.MS_LISTSWS 
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with valid or invalid parameters.
    /// <list type="bullet">
    ///     <item>CreateContentType</item>
    ///     <item>DeleteContentType</item>
    ///     <item>DeleteContentTypeXmlDocument</item>
    ///     <item>GetListContentType</item>
    ///     <item>UpdateContentType</item>
    ///     <item>GetListContentTypesAndProperties</item>
    ///     <item>UpdateContentTypeXMLDocument</item>
    ///     <item>UpdateContentTypesXmlDocument</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S02_OperationOnContentType : TestClassBase
    {
        #region Private member variables

        /// <summary>
        /// Protocol adapter
        /// </summary>
        private IMS_LISTSWSAdapter listswsAdapter;
        #endregion

        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">A instance of the testContext.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases

        #region ApplyContentTypeToList

        /// <summary>
        /// This test case is used to verify that the server returns soap fault without error code in ApplyContentTypeToList operation with not findable contentTypeID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC01_ApplyContentTypeToList_IncorrectContentTypeID()
        {
            string errorCodeForNullContentType = string.Empty;
            string errorCodeForNonExistentContentType = string.Empty;
            bool isExistentFaultForNullContentType = false;
            bool isExistentFaultForNonExistentContentType = false;
            bool isListCreatedSuccessfully = false;

            #region Add a list

            string listId = TestSuiteHelper.CreateList();
            isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region ApplyContentTypeToList

            // Call ApplyContentTypeToList method with valid content type Id.
            string contentTypeID = AdapterHelper.SiteDocumentContentTypeId;

            // Use list title to call ApplyContentTypeToList operation.
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyResult = this.listswsAdapter.ApplyContentTypeToList(null, contentTypeID, listId);
            this.Site.Assert.IsNotNull(
                        applyResult.Success,
                        "Apply content type {0} to the list {1} should succeed.",
                        contentTypeID,
                        listId);

            // Call ApplyContentTypeToList with invalid content type Id.
            contentTypeID = null;
            try
            {
                // Use a null content type
                this.listswsAdapter.ApplyContentTypeToList(null, contentTypeID, listId);
            }
            catch (SoapException exp)
            {
                errorCodeForNullContentType = TestSuiteHelper.GetErrorCode(exp);
                isExistentFaultForNullContentType = true;
            }

            contentTypeID = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            try
            {
                // Use a nonexistent content type
                this.listswsAdapter.ApplyContentTypeToList(null, contentTypeID, listId);
            }
            catch (SoapException exp)
            {
                errorCodeForNonExistentContentType = TestSuiteHelper.GetErrorCode(exp);
                isExistentFaultForNonExistentContentType = true;
            }

            // If the SOAP fault is returned, then capture R373.
            Site.CaptureRequirementIfIsTrue(
                isExistentFaultForNullContentType && isExistentFaultForNonExistentContentType,
                373,
                @"[In ApplyContentTypeToList operation] If the content type cannot be found, the protocol server MUST return a SOAP fault.");

            this.Site.Assert.IsNull(errorCodeForNullContentType, "There should be no error code for ApplyContentTypeToList operation when the content type cannot be found.");

            // If the returned error code is null, then capture R1625.
            Site.CaptureRequirementIfIsNull(
                errorCodeForNonExistentContentType,
                1625,
                @"[In ApplyContentTypeToList operation] [If the content type cannot be found, the protocol server MUST return a SOAP fault.]There is no error code returned for this fault.");

            #endregion test ApplyContentTypeToList
        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault with error code in ApplyContentTypeToList operation with incorrect list name which does not correspond to any lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC02_ApplyContentTypeToList_InvalidListName()
        {
            // Add a list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);
            this.Site.Assert.IsNotNull(listId, "Test suite should create the list successfully.");

            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string errorCode = string.Empty;
            bool isExistentFault = false;

            #region test ApplyContentTypeToList

            // Call ApplyContentTypeToList method with valid list name.
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyResult = this.listswsAdapter.ApplyContentTypeToList(null, AdapterHelper.SiteDocumentContentTypeId, listName);
            this.Site.Assert.IsNotNull(
                        applyResult.Success,
                        "Apply content type {0} to the list {1} should succeed.",
                        AdapterHelper.SiteDocumentContentTypeId,
                        listName);

            // Use a nonexistent list name
            try
            {
                this.listswsAdapter.ApplyContentTypeToList(null, AdapterHelper.SiteDocumentContentTypeId, invalidListName);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while ApplyContentTypeToList operation when the list name is invalid.");

            // If a SOAP fault is returned, then capture R385.
            Site.CaptureRequirementIfIsTrue(
                isExistentFault,
                385,
                @"[In ApplyContentTypeToList operation] If the operation fails, a SOAP fault MUST be returned.");

            if (Common.IsRequirementEnabled(2752, this.Site))
            {
                // If the error code is "0x82000006", then capture R2752.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    errorCode,
                    2752,
                    @"[In ApplyContentTypeToList operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");
            }
            #endregion test ApplyContentTypeToList
        }

        /// <summary>
        /// This test case is used to verify the ApplyContentTypeToList operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC03_ApplyContentTypeToList_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2433, this.Site), @"Test is executed only when R2433Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                #region Delete a content Type Xml Document with an invalid listname.

                // Call ApplyContentTypeToList method to apply an existing site content type to 
                // the requested list.
                this.listswsAdapter.ApplyContentTypeToList(null, string.Empty, invalidListName);

                #endregion
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2433
                bool isVerifyR2433 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2433 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2433 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2433,
                    2433,
                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<30> Section 3.1.4.6: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to test the ApplyContentTypeToList operation with correct contentTypeID and listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC04_ApplyContentTypeToList_WithListGuid_Succeed()
        {
            string contentTypeId = AdapterHelper.SiteDocumentContentTypeId;

            // Add a new list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Get all the current content types in the list.
            GetListContentTypesResponseGetListContentTypesResult result = this.listswsAdapter.GetListContentTypes(listId, null);
            this.Site.Assert.IsNotNull(result, "GetListContentTypes operation should succeed.");
            int numberBeforeApplyDocumentContentType = result.ContentTypes.ContentType.Length;

            // Use list title to call ApplyContentTypeToList operation.
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyResult = this.listswsAdapter.ApplyContentTypeToList(string.Empty, contentTypeId, listName);
            this.Site.Assert.IsNotNull(applyResult, "The response of operation ApplyContentTypeToList is NULL.");
            this.Site.Assert.IsNotNull(
                        applyResult.Success,
                        "Apply content type {0} to the list {1} should succeed.",
                        contentTypeId,
                        listId);

            if (Common.IsRequirementEnabled(16231, this.Site))
            {
                // Verify requirement R16231.
                // If there is no assert fail, it means ApplyContentTypeToList operation is successful, R16231 can be captured. 
                Site.CaptureRequirement(
                    16231,
                    @"Implementation does support this method[ApplyContentTypeToList]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // Apply the contentType to the new added list.
            result = this.listswsAdapter.GetListContentTypes(listId, null);
            this.Site.Assert.IsNotNull(result, "GetListContentTypes operation should succeed.");
            int numberAfterApplyDocumentContentType = result.ContentTypes.ContentType.Length;

            // If apply a new content type, then capture R376 and R377.
            Site.CaptureRequirementIfAreEqual<int>(
                numberBeforeApplyDocumentContentType + 1,
                numberAfterApplyDocumentContentType,
               376,
               @"[In ApplyContentTypeToList operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfAreEqual<int>(
                 numberBeforeApplyDocumentContentType + 1,
                 numberAfterApplyDocumentContentType,
                377,
                @"[In ApplyContentTypeToList operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            // Create a new list again.
            string secListId = TestSuiteHelper.CreateList();

            // Get all the current content types in the list.
            result = this.listswsAdapter.GetListContentTypes(secListId, null);
            this.Site.Assert.IsNotNull(result, "GetListContentTypes operation should succeed.");
            numberBeforeApplyDocumentContentType = result.ContentTypes.ContentType.Length;

            // Apply the document content type to the customer list.
            // The customer list will not contain document content type at first.
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyResultSecond = this.listswsAdapter.ApplyContentTypeToList(string.Empty, contentTypeId, secListId);
            this.Site.Assert.IsNotNull(applyResultSecond, "The response of operation ApplyContentTypeToList is NULL.");
            this.Site.Assert.IsNotNull(
                        applyResultSecond.Success,
                        "Apply content type {0} to the list {1} should succeed.",
                        contentTypeId,
                        secListId);

            // Apply the contentType to the new added list.
            result = this.listswsAdapter.GetListContentTypes(secListId, null);
            this.Site.Assert.IsNotNull(result, "GetListContentTypes operation should succeed.");
            numberAfterApplyDocumentContentType = result.ContentTypes.ContentType.Length;

            // If the result is successful, then directly capture the requirement.
            Site.CaptureRequirementIfAreEqual<int>(
                numberBeforeApplyDocumentContentType + 1,
                numberAfterApplyDocumentContentType,
                379,
                @"[In ApplyContentTypeToList operation] Otherwise [if the specified listName is a valid GUID and corresponds to the identification of a list on the site and the contentTypeId is a valid ContentTypeId corresponds to the identification of the content type that is to be associated with the list], after the method runs, the requested list will have the content type added to its list of valid content types.");
        }

        /// <summary>
        /// This test case is used to test the ApplyContentTypeToList operation with correct contentTypeID and listName GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC05_ApplyContentTypeToList_WithListTitle_Succeed()
        {
            #region Add a list

            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = false;
            isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            // Create another list and append the content type to the list.
            string newListId = TestSuiteHelper.CreateList();

            #region ApplyContentTypeToList

            // Apply the contentType to the list.
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyResult = this.listswsAdapter.ApplyContentTypeToList(string.Empty, AdapterHelper.SiteDocumentContentTypeId, newListId);
            this.Site.Assert.IsNotNull(applyResult, "The response of operation ApplyContentTypeToList is NULL.");

            // If the ApplyContentTypeToList is successful, then capture R374.
            Site.CaptureRequirementIfIsNotNull(
                applyResult.Success,
                374,
                @"[In ApplyContentTypeToList operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            #endregion ApplyContentTypeToList
        }

        #endregion

        #region CreateContentType

        /// <summary>
        /// This test case is used to test the result doesn't contain GetListContentTypesResult.ContentTypes.ContentTypeOrder in GetListContentTypes operation under Windows SharePoint Services 3.0.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC06_CreateContentType_ContentTypeOrderNotReturned_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2462, this.Site), @"Test is executed only when R2462Enabled is set to true.");

            // Initialize a valid ListName.
            string listname = TestSuiteHelper.CreateList();
            List<string> listString = new List<string>() { Common.GetConfigurationPropertyValue("ListFieldText", this.Site) };

            string contentTypeId = TestSuiteHelper.CreateContentType(
                                                                    listname,
                                                                    "CreateContentTypeDescriptionOf" + TestSuiteHelper.GenerateRandomString(5),
                                                                    listString);
            Site.Assert.IsNotNull(contentTypeId, "Create list success.");

            // Call GetListContentTypes method to get all content types from a list.
            GetListContentTypesResponseGetListContentTypesResult getListContentTypesResponseResult = this.listswsAdapter.GetListContentTypes(listname, contentTypeId);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2462
            // If the attribute GetListContentTypesResult.ContentTypes.ContentTypeOrder is not returned, 
            // then R2462 should be covered.
            Site.CaptureRequirementIfIsNull(
                getListContentTypesResponseResult.ContentTypes.ContentTypeOrder,
                2462,
                @"[In Appendix B: Product Behavior][GetListContentTypesResponse]Implementation does not return this attribute[GetListContentTypesResult.ContentTypes.ContentTypeOrder].(<52> Section 3.1.4.19.2.2: The attribute is not returned in Windows SharePoint Services 3.0.)");
        }

        /// <summary>
        /// This test case is used to verify that the server returns SOAP fault if the fields element contains invalid field name in CreateContentType operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC07_CreateContentType_ErrorFieldName()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "Creating a new list '{0}' should succeed.", listName);
            #endregion

            #region CreateContentType with error field name
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldNameB = TestSuiteHelper.GetUniqueFieldName();
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(fieldNameA, fieldNameB);
            string parentTypeName = TestSuiteHelper.GetFirstExistContentTypeName(listGuid);
            string parentTypeId = TestSuiteHelper.GetContentTypeId(listGuid, parentTypeName);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            // Call CreateContentType operation to create a content type with 2 field names, one is valid field name and the other is invalid one.
            contentTypeId = this.listswsAdapter.CreateContentType(
                                                        listGuid,
                                                        contentTypeDisplayName,
                                                        parentTypeId,
                                                        addFields,
                                                        new CreateContentTypeContentTypeProperties(),
                                                        "TRUE");

            Site.Assert.IsNotNull(contentTypeId, "The result of CreateContentType operation should not be NULL");

            #endregion

            #region Get the created content type and capture requirements.

            // Get the previous step created content type.
            GetListContentTypeResponseGetListContentTypeResult contentTypeResult = null;
            bool isExceptionThrown = false;

            try
            {
                contentTypeResult = this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);
            }
            catch (SoapException)
            {
                isExceptionThrown = true;
            }
            catch (WebException)
            {
                isExceptionThrown = true;
            }

            // If no exception thrown, it means the error is ignored by protocol server. Capture R429
            Site.CaptureRequirementIfIsFalse(
                        isExceptionThrown,
                        429,
                        @"The protocol server adds the given fields to the newly created content type. If an error regarding the fields is encountered during this operation, the protocol server MUST ignore this error.");

            bool isContains = false;
            foreach (FieldDefinitionCTRemove field in contentTypeResult.ContentType.Fields.Items)
            {
                if (field.Name == fieldNameA)
                {
                    isContains = true;
                    break;
                }
            }

            // If the content type exists, and the valid field exists, then capture R2266 and R2267.
            Site.CaptureRequirementIfIsTrue(
                isContains,
                2266,
                @"[The content type will still exist as a newly created content type on the protocol server and all fields added to the content type before the error was encountered MUST stay added to the content type.] The content type will still exist as a newly created content type on the protocol server.");

            Site.CaptureRequirementIfIsTrue(
                isContains,
                2267,
                @"[The protocol server adds the given fields to the newly created content type. If an error regarding the fields is encountered during this operation, the protocol server MUST ignore this error. The content type will still exist as a newly created content type on the protocol server and]  all fields added to the content type before the error was encountered MUST stay added to the content type.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify that the server returns a SOAP fault when there is property naming error with the content type still added in CreateContentType operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC08_CreateContentType_ErrorPropertyName()
        {
            #region Create a new list
            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "Creating a new list '{0}' should succeed.", listName);
            #endregion

            #region CreateContentType with error property name
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(fieldName);
            string parentTypeName = TestSuiteHelper.GetFirstExistContentTypeName(listGuid);
            string parentTypeId = TestSuiteHelper.GetContentTypeId(listGuid, parentTypeName);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            // Create a property naming issue by setting Title to a name ending with a period.
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Title = TestSuiteHelper.GetUniqueContentTypeName() + ".";
            bool soapFaultReturned = false;
            try
            {
                // Call CreateContentType operation to create a content type with invalid property name.
                this.listswsAdapter.CreateContentType(
                                                listGuid,
                                                contentTypeDisplayName,
                                                parentTypeId,
                                                addFields,
                                                addProperties,
                                                "TRUE");
            }
            catch (SoapException)
            {
                soapFaultReturned = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R430. SOAP fault is{0} returned after invoking CreateContentType operation with invalid property name.", soapFaultReturned ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                soapFaultReturned,
                430,
                @"The protocol server updates the content type with the given properties. If an error regarding the naming of properties, similar to the error specified in section 3.1.4.9.2.1 is encountered during this operation, the protocol server MUST return a SOAP fault.");
            #endregion

            #region Get the created content type and capture requirements.
            // Get the content type created in previous step.
            GetListContentTypesResponseGetListContentTypesResult contentTypesResult = null;
            contentTypesResult = this.listswsAdapter.GetListContentTypes(listName, null);

            GetListContentTypesResponseGetListContentTypesResultContentTypesContentType createdContentType = null;
            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType contentType in contentTypesResult.ContentTypes.ContentType)
            {
                if (contentType.Name == contentTypeDisplayName)
                {
                    createdContentType = contentType;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R2268.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2268
            // If the content type with the specified display name does exist, R2268 can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createdContentType,
                2268,
                @"The protocol server updates the content type with the given properties. If an error regarding the naming of properties, similar to the error specified in section 3.1.4.9.2.1 is encountered during this operation, [the protocol server MUST return a SOAP fault, and] the content type will still exist as a newly created content type on the protocol server.");

            // Get fields in the content type.
            Site.Assert.IsNotNull(createdContentType.ID, "Newly created content type '{0}' should contain identifier in GetListContentTypes response.", contentTypeDisplayName);
            GetListContentTypeResponseGetListContentTypeResult contentTypeResult = null;
            contentTypeResult = this.listswsAdapter.GetListContentType(listName, createdContentType.ID);

            bool fieldExists = false;
            foreach (FieldDefinitionCTRemove field in contentTypeResult.ContentType.Fields.Items)
            {
                if (field.Name == fieldName)
                {
                    fieldExists = true;
                    break;
                }
            }

            // If the field exists in the created content type, then capture R2269.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R2269. Field '{0}' is{1} added in content type '{2}'.", fieldName, fieldExists ? string.Empty : " not", contentTypeDisplayName);
            Site.CaptureRequirementIfIsTrue(
                fieldExists,
                2269,
                @"The protocol server updates the content type with the given properties. If an error regarding the naming of properties, similar to the error specified in section 3.1.4.9.2.1 is encountered during this operation, [the protocol server MUST return a SOAP fault, and the content type will still exist as a newly created content type on the protocol server.] All fields that were added MUST stay added to the content type.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault in CreateContentType operation with incorrect display name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC09_CreateContentType_IncorrectDisplayName()
        {
            #region Add a list
            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            string errorCode = string.Empty;
            bool isExistentFault = false;

            // Use a display name longer than 128 characters
            string displayName = "ContentTypeDisplayName" + TestSuiteHelper.GenerateRandomString(128);
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            string parentTypeID = TestSuiteHelper.GetContentTypeId(listId, TestSuiteHelper.GetFirstExistContentTypeName(listId));
            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNames.ToArray());

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = "CreateContentTypeDescriptionOf" + TestSuiteHelper.GenerateRandomString(5);
            addProperties.ContentType.Title = "CreateContentTitle" + TestSuiteHelper.GenerateRandomString(5);

            try
            {
                this.listswsAdapter.CreateContentType(
                                                    listId,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "SOAP fault should be returned for CreateContentType operation when the displayName is longer than 128 characters.");

            // If the returned error code is "0x80131600", then capture R436.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x80131600",
                errorCode,
                436,
                @"[In CreateContentType operation] [In CreateContentType element] [In displayName field] [A SOAP fault with error code 0x80131600 MUST be returned if] It [displayName] is longer than 128 characters.");

            isExistentFault = false;
            errorCode = string.Empty;

            // Use a display name ending with period
            displayName = "CreateContentTypeDisplayName" + TestSuiteHelper.GenerateRandomString(5) + ".";

            try
            {
                this.listswsAdapter.CreateContentType(
                                                    listId,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "SOAP fault should be returned for CreateContentType operation when the displayName ends with a period.");
            Site.CaptureRequirementIfAreEqual<string>(
                "0x80131600",
                errorCode,
                437,
                @"[In CreateContentType operation] [In CreateContentType element] [In displayName field] [A SOAP fault with error code 0x80131600 MUST be returned if] It [displayName] ends with a period.");

            isExistentFault = false;
            errorCode = string.Empty;

            // Use a display name ending with whitespace
            displayName = "ContentTypeDisplayName" + TestSuiteHelper.GenerateRandomString(5) + " ";

            try
            {
                this.listswsAdapter.CreateContentType(
                                                    listId,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting CreateContentType operation with whitespace at the end in DisplayName.");
            Site.CaptureRequirementIfAreEqual<string>(
                "0x80131600",
                errorCode,
                4837,
                @"[In CreateContentType operation] [In CreateContentType element] [In displayName field] [A SOAP fault with error code 0x80131600 MUST be returned if] It [displayName] begins or ends with whitespace.");

            isExistentFault = false;
            errorCode = string.Empty;

            // Verify R440 and R428
            const string ValidDisplayName = "ValidDisplayName";
            displayName = string.Format("{0}", ValidDisplayName);
            string contentTypeId = string.Empty;

            try
            {
                contentTypeId = this.listswsAdapter.CreateContentType(
                                                                listId,
                                                                displayName,
                                                                parentTypeID,
                                                                addField,
                                                                addProperties,
                                                                "TRUE");
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsFalse(isExistentFault, "The CreateContentType operation with valid display name '{0}' should be successful.", displayName);

            isExistentFault = false;
            errorCode = string.Empty;

            try
            {
                this.listswsAdapter.DeleteContentType(listId, contentTypeId);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsFalse(isExistentFault, "The DeleteContentType operation with content type Id '{0}' should be successful.", contentTypeId);

            isExistentFault = false;
            errorCode = string.Empty;

            // A container which contains all the invalid chars.
            string[] invalidLetterList = new string[] { "\t", "..", "*", "?", "\"", "#", "%", "<", ">", "{", "}", "|", "~", "&" };

            foreach (string invalidLetter in invalidLetterList)
            {
                // Use a display name containing invalid characters
                // The invalid display name format is: the valid display name appends an invalid letter.
                displayName = string.Format(
                                        "{0}{1}",
                                        ValidDisplayName,
                                        invalidLetter);

                try
                {
                    this.listswsAdapter.CreateContentType(
                                                        listId,
                                                        displayName,
                                                        parentTypeID,
                                                        addField,
                                                        addProperties,
                                                        "TRUE");
                }
                catch (SoapException exp)
                {
                    errorCode = TestSuiteHelper.GetErrorCode(exp);
                    isExistentFault = true;
                }

                this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting CreateContentType operation when the displayName is {0}", displayName);

                this.Site.Assert.AreEqual<string>(
                    "0x80131600",
                    errorCode,
                    "SOAP fault code should be 0x80131600 for CreateContentType operation when the displayName is {0}",
                    displayName);
            }

            // If CreateContentType operation is successful for valid display name and 
            // returns SoapFault and error code 0x80131600 for all the invalid display names,
            // then capture the requirement R440.
            Site.CaptureRequirement(
                440,
                @"[In CreateContentType operation] [In CreateContentType element] [In displayName field] [A SOAP fault with error code 0x80131600 MUST be returned if] It [displayName] contains a tab, two consecutive periods, or any of the following characters: \/: * ? "" # % < > { } | ~ &.");

            // If CreateContentType operation is successful for valid display name and 
            // returns SoapFault and error code 0x80131600 for all the invalid display names,
            // then capture the requirement R428.
            Site.CaptureRequirement(
                428,
                @"[In CreateContentType operation] The protocol server attempts to create a new content type on the list with the given parent content type and display name. If an error specified in section 3.1.4.9.2.1 is encountered during this operation, the protocol server MUST return a SOAP fault.");

            #endregion test CreateContentType

            

        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault with error code in CreateContentType operation with invalid ListName which does not correspond to any lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC10_CreateContentType_InvalidListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2753, this.Site), @"Test is executed only when R2753Enabled is set to true.");

            // Create a new listName
            string listName = TestSuiteHelper.GetUniqueListName();

            #region CreateContentType

            // Create a contentType with invalid listName.
            string errorCode = string.Empty;
            bool isExistentFault = false;
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);

            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNames.ToArray());

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = displayName;
            addProperties.ContentType.Title = displayName;

            try
            {
                this.listswsAdapter.CreateContentType(
                                                    listName,
                                                    displayName,
                                                    null,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {
                isExistentFault = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting CreateContentType operation.");

            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2753,
                @"[In CreateContentType operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");

            #endregion test CreateContentType
        }

        /// <summary>
        /// This test case is used to verify the CreateContentType operation in Windows SharePoint Services 3.0 
        /// when ListName parameter can't be found in current existed lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC11_CreateContentType_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2438, this.Site), @"Test is executed only when R2438Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                #region Create a content type with invalid ListName on the list

                // Call CreateContentType method to create a content type with an invalid ListName.
                this.listswsAdapter.CreateContentType(
                                                        invalidListName,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                #endregion
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2438
                bool isVerifyR2438 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2438 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2438 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2438,
                    2438,
                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<37> Section 3.1.4.9: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault in CreateContentType operation with incorrect parentType.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC12_CreateContentType_InvalidParentType()
        {
            #region Add a list

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            string errorCode = string.Empty;
            bool isExistentFault = false;
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);

            // Invalid parentTypeID
            string parentTypeID = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNames.ToArray());

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = displayName;
            addProperties.ContentType.Title = displayName;

            try
            {
                this.listswsAdapter.CreateContentType(
                                                    listId,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting CreateContentType operation when the parent content type is not valid.");

            // If SOAP fault is returned, then capture R426 and R427.
            Site.CaptureRequirementIfIsTrue(
                isExistentFault,
                426,
                @"[In CreateContentType operation] If the parentType is not a valid content type identifier, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                isExistentFault,
                427,
                @"[In CreateContentType operation] If the parentType is not the content type identifier of a content type available on the list, the protocol server MUST return a SOAP fault.");

            // If no error codes is returned, then capture R1683.
            Site.CaptureRequirementIfIsNull(
                errorCode,
                1683,
                @"[In CreateContentType operation] [If the parentType is not a valid content type identifier, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            #endregion test CreateContentType
        }

        /// <summary>
        /// This test case is used to test the CreateContentType operation with correct list name GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC13_CreateContentType_UseListGuid_Succeed()
        {
            #region Add a list

            // Add a list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType with the listName which is a valid GUID.
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            string parentTypeID = TestSuiteHelper.GetContentTypeId(listId, TestSuiteHelper.GetFirstExistContentTypeName(listId));
            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNames.ToArray());

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = displayName;
            addProperties.ContentType.Title = displayName;

            bool createContentTypeSucceeded = false;
            string contentType = this.listswsAdapter.CreateContentType(
                                                    listId,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);

            // Verify requirement R16801.
            // If there are no other errors, it means implementation does support this CreateContentType method. R16801 can be captured.
            if (Common.IsRequirementEnabled(16801, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                    createContentTypeSucceeded,
                    16801,
                    @"Implementation does support this method[CreateContentType]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If the CreateContentType is successful, then capture R421.
            Site.CaptureRequirementIfIsTrue(
                createContentTypeSucceeded,
                421,
                @"[In CreateContentType operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            #endregion test CreateContentType
        }

        /// <summary>
        /// This test case is used to test the CreateContentType operation with correct list name title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC14_CreateContentType_UseListTitle_Succeed()
        {
            #region Add a list

            // Get a listName which is an invalid GUID.
            string listName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList(listName);
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType with a listName which is an invalid GUID.
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            string parentTypeID = TestSuiteHelper.GetContentTypeId(listId, TestSuiteHelper.GetFirstExistContentTypeName(listId));
            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNames.ToArray());

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = displayName;
            addProperties.ContentType.Title = displayName;

            bool createContentTypeSucceeded = false;
            string contentType = this.listswsAdapter.CreateContentType(
                                                    listName,
                                                    displayName,
                                                    parentTypeID,
                                                    addField,
                                                    addProperties,
                                                    "TRUE");
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);

            // If the CreateContentType is successful, then capture R422.
            Site.CaptureRequirementIfIsTrue(
                createContentTypeSucceeded,
                422,
                @"[In CreateContentType operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion test CreateContentType
        }

        #endregion

        #region DeleteContentTypeXmlDocument

        /// <summary>
        /// This test case is used to verify the DeleteContentTypeXmlDocument operation returning the SOAP Fault if the content type is specified by the TypeId.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC15_DeleteContentTypeXmlDocument_ContentTypeCannotBeFound()
        {
            #region Invoke AddList operation to create two generic list (one is List_1, another is List_2).
            string slist_GUID_1 = TestSuiteHelper.GetUniqueListName();
            string slist_GUID_2 = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(slist_GUID_1);
            TestSuiteHelper.CreateList(slist_GUID_2);
            #endregion

            #region Invoke CreateContentType operation to create new content type for each generic list.(contentType1 for List_1, contentType2 is List_2).
            string contentTypeId_ForList1 = null;
            string contentTypeId_ForList2 = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName1 = TestSuiteHelper.GetUniqueContentTypeName();
            string contentTypeDisplayName2 = TestSuiteHelper.GetUniqueContentTypeName();
            contentTypeId_ForList1 = TestSuiteHelper.CreateContentType(slist_GUID_1, contentTypeDisplayName1, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId_ForList1, "Error: Failed to create a new content type by CreateContentType operation!");

            contentTypeId_ForList2 = TestSuiteHelper.CreateContentType(slist_GUID_2, contentTypeDisplayName2, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId_ForList1, "Error: Failed to create another new content type by CreateContentType operation!");
            #endregion

            #region Invoke UpdateContentTypeXmlDocument operation to append a new XML document for the contentType1 in List1.
            string sxmlDocUrl = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            string elementName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, sxmlDocUrl, innerXml);
            this.listswsAdapter.UpdateContentTypeXmlDocument(slist_GUID_1, contentTypeId_ForList1, newDocument);
            #endregion

            #region Invoke DeleteContentTypeXmlDocument operation in the generic list List1 with the new XML Document, but set the value of "contentTypeId" to contentType2.
            // Invoke DeleteContentTypeXmlDocument operation in the generic list "List1" with the new XML Document, 
            // but set the value of "contentTypeId" to contentType2.
            // If a SOAP fault is obtained in the response of DeleteContentTypeXmlDocument, R506 is captured.
            // If a SOAP fault is obtained without error code in the response of DeleteContentTypeXmlDocument, R1726 is captured.
            bool isSoapFaultExisting = false;
            string errorCode = string.Empty;
            try
            {
                this.listswsAdapter.DeleteContentTypeXmlDocument(slist_GUID_1, contentTypeId_ForList2, sxmlDocUrl);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisting = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            // if there is a soap fault and no error code, capture R506 and R1726
            this.Site.Log.Add(
                            LogEntryKind.Debug,
                            "Capture requirement :#R506 and #R1726 IsSoapFaultExistring:[{0}],current error code[{1}].",
                            isSoapFaultExisting,
                            string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultExisting && string.IsNullOrEmpty(errorCode),
                506,
                @"[In DeleteContentTypeXmlDocument operation] If the content type specified by the TypeId element "
                + "is not found, a SOAP fault MUST be thrown.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultExisting && string.IsNullOrEmpty(errorCode),
                1726,
                @"[In DeleteContentTypeXmlDocument operation] [If the content type specified by the TypeId element "
                + "is not found, a SOAP fault MUST be thrown.]There is no error code with this fault.");
        }

        /// <summary>
        /// This test case is used to test the successful status of DeleteContentTypeXmlDocument operation with an invalid GUID listName which corresponds to the list title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC16_DeleteContentTypeXmlDocument_InvalidListName()
        {
            #region Add a list

            // Add a new list to the server.
            // Create a listName which is an invalid GUID.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument

            string elementName = Common.GetConfigurationPropertyValue("ContentTypeXmlQualifiedName", this.Site);
            string nameSpaceUri = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innnerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            System.Xml.XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, nameSpaceUri, innnerXml);
            bool isNewDocumentCreatedSuccessfully = !string.IsNullOrEmpty(newDocument.ToString());
            this.Site.Assert.IsTrue(isNewDocumentCreatedSuccessfully, "New document should be created successfully.");

            // Update the XmlDocument
            System.Xml.XPath.IXPathNavigable updateResult = this.listswsAdapter.UpdateContentTypeXmlDocument(listId, contentType, newDocument);
            this.Site.Assert.IsNotNull(updateResult, "UpdateContentTypeXmlDocument operation should succeed.");

            #endregion

            #region  DeleteContentTypeXmlDocument

            DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult deleteResult = this.listswsAdapter.DeleteContentTypeXmlDocument(listName, contentType, nameSpaceUri);

            // If the list whose list title corresponds to the specified listName exists, then DeleteContentTypeXmlDocument operation will succeed. The "Success" element is returned only when it succeeds.
            // If the Success element isn't null, then capture R503 and R504.
            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                503,
                @"[In DeleteContentTypeXmlDocument operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                504,
                @"If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion DeleteContentTypeXmlDocument
        }

        /// <summary>
        /// This test case is used to verify the DeleteContentTypeXmlDocument operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC17_DeleteContentTypeXmlDocument_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2446, this.Site), @"Test is executed only when R2446Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                #region Delete a content Type Xml Document with an invalid listname.

                // Call DeleteContentTypeXmlDocument method to delete a content type with an invalid ListName.
                this.listswsAdapter.DeleteContentTypeXmlDocument(invalidListName, null, null);

                #endregion
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2446
                bool isVerifyR2446 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2446 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2446 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2446,
                    2446,
                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<42> Section 3.1.4.12: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault with error code in DeleteContentTypeXmlDocument operation with non-existing listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC18_DeleteContentTypeXmlDocument_NonExistentListName()
        {
            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument

            string elementName = Common.GetConfigurationPropertyValue("ContentTypeXmlQualifiedName", this.Site);
            string nameSpaceUri = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innnerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            System.Xml.XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, nameSpaceUri, innnerXml);
            bool isNewDocumentCreatedSuccessfully = newDocument != null;
            this.Site.Assert.IsTrue(isNewDocumentCreatedSuccessfully, "CreateNewDocument should succeed.");

            // Update the XmlDocument
            System.Xml.XPath.IXPathNavigable updateResult = this.listswsAdapter.UpdateContentTypeXmlDocument(listId, contentType, newDocument);
            this.Site.Assert.IsNotNull(updateResult, "UpdateContentTypeXmlDocument operation should succeed.");

            #endregion

            #region  DeleteContentTypeXmlDocument

            string errorCode = string.Empty;
            string nonexistListName = TestSuiteHelper.GetUniqueListName();
            bool isExistentFault = false;

            try
            {
                this.listswsAdapter.DeleteContentTypeXmlDocument(nonexistListName, contentType, nameSpaceUri);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault for DeleteContentTypeXmlDocument with the listName does not correspond to any lists.");

            // If a SOAP fault is returned, then capture R514.
            Site.CaptureRequirementIfIsTrue(
                isExistentFault,
                514,
                @"[In DeleteContentTypeXmlDocument operation] If the operation fails, a SOAP fault MUST be returned.");

            if (Common.IsRequirementEnabled(2756, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    errorCode,
                    2756,
                    @"[In DeleteContentTypeXmlDocument operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");
            }

            bool isNonExistentList = false;

            try
            {
                this.listswsAdapter.GetList(nonexistListName);
            }
            catch (SoapException)
            {
                isNonExistentList = true;
            }

            // If the list does not exist or has been deleted by another user, GetList operation will return a soap fault.
            Site.CaptureRequirementIfIsTrue(
                isNonExistentList,
                1725,
                @"[In DeleteContentTypeXmlDocument operation] [If the specified listName does not correspond to a list from either of these checks, the protocol server should<40> return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");

            #endregion DeleteContentTypeXmlDocument
        }

        /// <summary>
        ///  This test case is used to test the DeleteContentTypeXmlDocument operation with correct listName GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC19_DeleteContentTypeXmlDocument_Succeed()
        {
            #region Add a list

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument

            string elementName = Common.GetConfigurationPropertyValue("ContentTypeXmlQualifiedName", this.Site);
            string nameSpaceUri = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innnerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            System.Xml.XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, nameSpaceUri, innnerXml);
            bool isNewDocumentCreatedSuccessfully = !string.IsNullOrEmpty(newDocument.ToString());
            this.Site.Assert.IsTrue(isNewDocumentCreatedSuccessfully, "New document should be created successfully.");

            // Update the XmlDocument
            System.Xml.XPath.IXPathNavigable updateResult = this.listswsAdapter.UpdateContentTypeXmlDocument(listId, contentType, newDocument);
            this.Site.Assert.IsNotNull(updateResult, "UpdateContentTypeXmlDocument operation should succeed.");

            #endregion

            #region  DeleteContentTypeXmlDocument

            DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult deleteResult = this.listswsAdapter.DeleteContentTypeXmlDocument(listId, contentType, nameSpaceUri);

            // Verify requirement R4971.
            // If delete is successful, it means implementation does support this DeleteContentTypeXmlDocument method. R4971 can be captured.
            if (Common.IsRequirementEnabled(4971, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    deleteResult.Success,
                    4971,
                    @"Implementation does support this method[DeleteContentTypeXmlDocument]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If the DeleteContentTypeXmlDocument succeed, then capture R502 and R508.
            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                502,
                @"[In DeleteContentTypeXmlDocument operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                deleteResult,
                508,
                @"[In DeleteContentTypeXmlDocument operation] If no SOAP fault has been thrown, the protocol server MUST respond with DeleteContentTypeXmlDocumentResult.");

            #endregion DeleteContentTypeXmlDocument
        }

        /// <summary>
        /// This test case is used to verify the successful status of DeleteContentTypeXmlDocument operation with deleting an xml document from content type.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC20_DeleteContentTypeXmlDocument_Success()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC19_DeleteContentTypeXmlDocument_Success, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            #region  UpdateContentTypeXMLDocument

            string namespaceUri1 = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            string elementName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Initialize a newDocument.
            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, namespaceUri1, innerXml);

            // Call UpdateContentTypeXmlDocument method to update XML Document properties of the content type collection.
            System.Xml.XPath.IXPathNavigable result = null;
            result = this.listswsAdapter.UpdateContentTypeXmlDocument(listName, contentTypeId, newDocument);
            Site.Assert.IsNotNull(result, "Call UpdateContentTypeXmlDocument operation successfully.");

            #endregion

            #region  DeleteContentTypeXmlDocument

            bool isCaptured = true;
            try
            {
                // Call DeleteContentTypeXmlDocument method to delete a content type xml document.
                this.listswsAdapter.DeleteContentTypeXmlDocument(listName, contentTypeId, namespaceUri1);
            }
            catch (SoapException)
            {
                Site.Assert.Fail("DeleteContentTypeXmlDocument failed.");
                isCaptured = false;
            }
            catch (WebException)
            {
                Site.Assert.Fail("DeleteContentTypeXmlDocument failed.");
                isCaptured = false;
            }

            // If the content type xml document is deleted successfully, then R507 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                507,
                @"[In DeleteContentTypeXmlDocument operation] If an XML document with the same namespace URI [RFC2396] specified by documentUri exists, it MUST be deleted.");

            #endregion DeleteContentTypeXmlDocument
        }

        #endregion

        #region DeleteContentType

        /// <summary>
        /// This test case is used to test operation DeleteContentType when the value of element "contentTypeId" cannot be found in the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC21_DeleteContentType_ContentTypeCannotBeFound()
        {
            #region Invoke AddList operation to create two generic list (one is List_1, another is List_2).
            string listGuid = TestSuiteHelper.GetUniqueListName();
            string listGuidTwo = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listGuid);
            TestSuiteHelper.CreateList(listGuidTwo);
            #endregion

            #region Invoke CreateContentType operation to create new content type for each generic list.(contentType1 for List_1, contentType2 is List_2).
            string contentTypeId_ForListOne = null;
            string contentTypeId_ForListTwo = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId_ForListOne = TestSuiteHelper.CreateContentType(listGuid, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId_ForListOne, "Error: Failed to create a new content type by CreateContentType operation!");

            contentTypeId_ForListTwo = TestSuiteHelper.CreateContentType(listGuidTwo, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId_ForListOne, "Error: Failed to create another new content type by CreateContentType operation!");
            #endregion

            #region Invoke DeleteContentType operation in the generic list List1, but set the value of "contentTypeId" to contentType2.
            // Invoke DeleteContentType operation in the generic list "List1", and set the value of "contentTypeId" to contentType2
            // so that the content type id cannot be found in List1.
            // If the product server returns a SOAP fault, R489 is captured.
            bool isR489Verified = false;
            try
            {
                this.listswsAdapter.DeleteContentType(listGuid, contentTypeId_ForListTwo);
            }
            catch (SoapException)
            {
                isR489Verified = true;
            }
            #endregion

            #region Capture requirement #489 when we get a SOAP fault in the response of DeleteContentType.
            Site.CaptureRequirementIfIsTrue(
                isR489Verified,
                489,
                @"[In DeleteContentType operation] If the content type cannot be found on the list, the protocol server MUST return a SOAP fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that the server returns soap fault with error code in DeleteContentType operation with incorrect listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC22_DeleteContentType_IncorrectListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2755, this.Site), @"Test is executed only when R2755Enabled is set to true.");

            #region Add a list

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  DeleteContentType

            // Create a new listName
            string listName = TestSuiteHelper.GetUniqueListName();
            string errorCode = string.Empty;
            bool isFaultExist = false;

            try
            {
                this.listswsAdapter.DeleteContentType(listName, contentType);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isFaultExist = true;
            }

            this.Site.Assert.IsTrue(isFaultExist, "The server response should contain a SOAP fault while requesting DeleteContentType when the listName does not correspond to List name or list title.");

            // If the returned error code is 0x82000006, then capture R2755.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2755,
                @"[In DeleteContentType operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");

            #endregion DeleteContentType
        }

        /// <summary>
        /// This test case is used to verify DeleteContentType operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC23_DeleteContentType_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2443, this.Site), @"Test is executed only when R2443Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;

            try
            {
                #region Delete a content type with invalid Listname.

                // Call DeleteContentType method to delete a content type with an invalid ListName.
                this.listswsAdapter.DeleteContentType(
                                                        invalidListName,
                                                        "0x00" + Guid.NewGuid().ToString("N"));

                #endregion
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2443
                bool isVerifyR2443 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2443 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2443 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2443,
                    2443,
                    @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<40> Section 3.1.4.11: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to verify that server returns soap fault without error code in DeleteContentType operation with non-existing.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC24_DeleteContentType_NonExistentContentType()
        {
            #region Add a list

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  DeleteContentType

            // Create an invalid contentTypeID
            string contentTypeID = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string errorCode = string.Empty;
            bool isExistentFault = false;

            try
            {
                this.listswsAdapter.DeleteContentType(listId, contentTypeID);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentFault = true;
            }

            this.Site.Assert.IsTrue(isExistentFault, "The server response should contain a SOAP fault while requesting DeleteContentType operation when the contentType cannot be found.");

            // If a SOAP fault is returned, then capture R496.
            Site.CaptureRequirementIfIsTrue(
                isExistentFault,
                496,
                @"[In DeleteContentType operation] If the operation fails, a SOAP fault MUST be returned.");

            // If the error code returned is null, then capture R1714.
            Site.CaptureRequirementIfIsNull(
                errorCode,
                1714,
                @"[In DeleteContentType operation] [If the content type cannot be found on the list, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            #endregion DeleteContentType
        }

        /// <summary>
        /// This test case is used to test the DeleteContentType operation with an invalid GUID listName but corresponds to the title of list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC25_DeleteContentType_UseListTitle_Succeed()
        {
            #region Add a list

            // Create a listName which is an invalid GUID.
            string listName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Add a new list to the server with the newly created listName.
            string listId = TestSuiteHelper.CreateList(listName);
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");
            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  DeleteContentType

            DeleteContentTypeResponseDeleteContentTypeResult deleteResult = this.listswsAdapter.DeleteContentType(listName, contentType);

            // Verify requirement R4801.
            // If delete is successful, it means implementation does support this DeleteContentType method. R4801 can be captured.
            if (Common.IsRequirementEnabled(4801, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    deleteResult.Success,
                    4801,
                    @"Implementation does support this method[DeleteContentType]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If DeleteContentType operation succeeds, then capture R486 and R487.
            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                486,
                @"[In DeleteContentType operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                487,
                @"[In DeleteContentType operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion DeleteContentType
        }

        /// <summary>
        /// This test case is used to test the DeleteContentType operation with correct listName with GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC26_DeleteContentType_WithListGuid_Succeed()
        {
            #region Add a list

            // Add a new list to the server.
            string listId = TestSuiteHelper.CreateList();
            bool isListCreatedSuccessfully = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessfully, "Test suite should create the list successfully.");

            #endregion Add a list

            #region CreateContentType

            // Create a contentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            List<string> fieldNames = new List<string>();
            string fieldsName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            fieldNames.Add(fieldsName);
            bool createContentTypeSucceeded = false;
            string contentType = TestSuiteHelper.CreateContentType(listId, displayName, fieldNames);
            createContentTypeSucceeded = !string.IsNullOrEmpty(contentType);
            this.Site.Assert.IsTrue(createContentTypeSucceeded, "CreateContentType Should succeed.");

            #endregion CreateContentType

            #region  DeleteContentType

            DeleteContentTypeResponseDeleteContentTypeResult deleteResult = this.listswsAdapter.DeleteContentType(listId, contentType);

            Site.CaptureRequirementIfIsNotNull(
                deleteResult.Success,
                485,
                @"[In DeleteContentType operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            #endregion DeleteContentType

            #region verify content type is deleted

            bool isExistentContentType = false;
            string errorCode = string.Empty;

            // If the specified contentType has been deleted, then GetListContentType operation will return a SOAP fault and there is no errorCode.
            try
            {
                this.listswsAdapter.GetListContentType(listId, contentType);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isExistentContentType = true;
            }

            Site.CaptureRequirementIfIsTrue(
                errorCode == null && isExistentContentType,
                490,
                @"[In DeleteContentType operation] If there are no other errors, the content type is removed from the list.");

            #endregion
        }

        #endregion

        #region GetListContentType

        /// <summary>
        /// This test case is used to test GetListContentType with incorrect contentTypeID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC27_GetListContentType_IncorrectContentType()
        {
            #region Add a list

            string listId = TestSuiteHelper.CreateList();

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list succeeded.");

            #endregion Add a list

            #region CreateContentType
            List<string> filedNames = new List<string>() { Common.GetConfigurationPropertyValue("ListFieldText", this.Site) };
            string parentTypeName = TestSuiteHelper.GetFirstExistContentTypeName(listId);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(parentTypeName), "Get exist Content Type succeeded.");

            string createdContentTypeID = TestSuiteHelper.CreateContentType(
                                                                        listId,
                                                                        "CreateContentTypeDescriptionOf" + TestSuiteHelper.GenerateRandomString(5),
                                                                        filedNames);

            bool isCreateContentTypeSuccessful = !string.IsNullOrEmpty(createdContentTypeID);
            this.Site.Assert.IsTrue(isCreateContentTypeSuccessful, "CreateContentType succeeded.");

            #endregion CreateContentType

            #region GetListContentType
            // Use incorrect contentTypeID, it should be invalid guid type format.
            string invalidContentTypeID = Guid.NewGuid().ToString("N");
            bool isExistSoapFault = false;
            string errorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetListContentType(listId, invalidContentTypeID);
            }
            catch (SoapException soapEx)
            {
                isExistSoapFault = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            Site.Assert.IsTrue(isExistSoapFault, "The server response should contain a SOAP fault while requesting GetListContentType when the specified ContentType id is not exist.");

            Site.CaptureRequirementIfIsTrue(
                isExistSoapFault,
                595,
                @"[In GetListContentType operation] If contentTypeId does not match a content type identifier on the list, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(errorCode),
                1805,
                @"[In GetListContentType operation] [If contentTypeId does not match a content type identifier on the list, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetListContentType with incorrect list name title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC28_GetListContentType_IncorrectListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2761, this.Site), @"Test is executed only when R2761Enabled is set to true.");

            #region Add a list

            string listId = TestSuiteHelper.CreateList();

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list succeeded.");

            #endregion Add a list

            #region CreateContentType
            List<string> filedNames = new List<string>() { Common.GetConfigurationPropertyValue("ListFieldText", this.Site) };
            string parentTypeName = TestSuiteHelper.GetFirstExistContentTypeName(listId);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(parentTypeName), "Getting exist ContentType successfully.");

            string createdContentTypeID = TestSuiteHelper.CreateContentType(
                                                                        listId,
                                                                        "CreateContentTypeDescriptionOf" + TestSuiteHelper.GenerateRandomString(5),
                                                                        filedNames);

            bool isCreateContentTypeSuccessful = !string.IsNullOrEmpty(createdContentTypeID);
            this.Site.Assert.IsTrue(isCreateContentTypeSuccessful, "Calling CreateContentType successfully.");

            #endregion CreateContentType

            #region GetListContentType

            string invalidList = DateTime.Now.TimeOfDay.ToString();
            bool caughtSoapException = false;
            try
            {
                this.listswsAdapter.GetListContentType(invalidList, string.Empty);
                Site.Assert.Fail("GetListContentType should fail when the list name is incorrect, but actually it passed.");
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    2761,
                    @"[In GetListContentType operation]Implementation does return a SOAP fault with error code 0x82000006,  if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    1804,
                    @"[In GetListContentType operation] [If the specified listName does not correspond to a list from either of these checks, the protocol server should<49> return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            }
            finally
            {
                Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetListContentType' with incorrect list name title.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetListContentType operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC29_GetListContentType_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2457, this.Site), @"Test is executed only when R2457Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                #region Delete a content Type Xml Document with an invalid listname.

                // Call GetListContentType method to get content type data for a given content type identifier.
                this.listswsAdapter.GetListContentType(invalidListName, string.Empty);

                #endregion
            }
            catch (SoapException soapException)
            {
                isReturnSopeFault = true;

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2457
                bool isVerifyR2457 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2457 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2457 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2457,
                    2457,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<49> Section 3.1.4.18: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to test operation GetListContentType when the value of element 
        /// listName is title of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC30_GetListContentType_UseListTitleAsListName()
        {
            #region Invoke AddList to create a generic list.
            string listGUID = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke CreateContentType operation to create new content type for the new generic list.
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();
            contentTypeId = TestSuiteHelper.CreateContentType(listGUID, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Error: Failed to create a new content type by CreateContentType operation!");
            #endregion

            #region Invoke GetList operation to get the title of the new generic list.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(listGUID);
            Site.Assert.IsNotNull(listDef, "Error: the object \"listDef\" is null !");
            Site.Assert.IsNotNull(listDef.Title, "Error: the object \"listDef.Title\" is null !");
            string listTitle = listDef.Title;
            #endregion

            #region Invoke GetListContentType operation with the new content type in the generic list, set the value of listName as the title of the list.
            GetListContentTypeResponseGetListContentTypeResult getListContentTypeResult = null;
            getListContentTypeResult = this.listswsAdapter.GetListContentType(listTitle, contentTypeId);

            // Verify requirement R5861.
            // If getting content type is successful, it means implementation does support this GetListContentType method. R5861 can be captured.
            if (Common.IsRequirementEnabled(5861, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    getListContentTypeResult,
                    5861,
                    @"Implementation does support this method[GetListContentType]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If we get the successful response of GetListContentType, we can capture R592 and R593.
            bool isR592Verified = false;
            bool isR593Verified = false;
            if (getListContentTypeResult != null)
            {
                isR592Verified = true;
                isR593Verified = true;
            }
            #endregion

            #region Capture requirements #592 and #593 when we get the successful response of GetListContentType.
            Site.CaptureRequirementIfIsTrue(
                isR592Verified,
                592,
                @"[In GetListContentType operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsTrue(
                isR593Verified,
                593,
                @"[In GetListContentType operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test operation GetListContentType when the value of element 
        /// listName is the GUID of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC31_GetListContentType_UseValidGuidAsListName()
        {
            #region Invoke AddList to create a generic list.
            string listGUID = TestSuiteHelper.CreateList();
            #endregion

            #region Invoke CreateContentType operation to create new content type for the new generic list.
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();
            contentTypeId = TestSuiteHelper.CreateContentType(listGUID, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Error: Failed to create a new content type by CreateContentType operation!");
            #endregion

            #region Invoke GetListContentType operation with the new content type in the generic list, set the value of listName as the GUID of the list.
            GetListContentTypeResponseGetListContentTypeResult getListContentTypeResult = null;
            getListContentTypeResult = this.listswsAdapter.GetListContentType(listGUID, contentTypeId);

            // If we get the successful response of GetListContentType, we can capture R591 and R596.
            bool isR591Verified = false;
            bool isR596Verified = false;
            if (getListContentTypeResult != null)
            {
                isR591Verified = true;
                isR596Verified = true;
            }
            #endregion

            #region  Capture requirements #591 and #596 when we get the successful response of GetListContentType.
            Site.CaptureRequirementIfIsTrue(
                isR591Verified,
                591,
                @"[In GetListContentType operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfIsTrue(
                isR596Verified,
                596,
                @"[In GetListContentType operation] If no other errors occurred the protocol server MUST respond with a GetListContentTypeResponse.");
            #endregion
        }

        #endregion

        #region GetListContentTypesAndProperties

        /// <summary>
        /// This test case is used to test server behavior when at least one of the input parameter is invalid in GetListContentTypesAndProperties operation.   
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC32_GetListContentTypesAndProperties_Fault()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Invoke GetListContentTypesAndProperties operation with listName invalid.
            string invalidListName = TestSuiteHelper.GetUniqueListName();

            bool isErrorCodeVerified = false;
            try
            {
                this.listswsAdapter.GetListContentTypesAndProperties(invalidListName, string.Empty, string.Empty, false, false);
            }
            catch (SoapException e)
            {
                string soapErrorCode = e.Detail.LastChild.InnerText;
                if (soapErrorCode == "0x82000006")
                {
                    isErrorCodeVerified = true;
                }
            }
            #endregion

            #region Capture R1042 R1043
            Site.CaptureRequirementIfIsTrue(
                isErrorCodeVerified,
                1042,
                @"[In GetListContentTypesAndProperties]If listName does not correspond to a list from either of these checks, the protocol server MUST return a SOAP fault with error code 0x82000006.");
            Site.CaptureRequirementIfIsTrue(
                isErrorCodeVerified,
                1043,
                @"[In [In GetListContentTypesAndProperties][If listName does not correspond to a list from either of these checks, the protocol server MUST return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the Version attribute in the response in GetListContentTypesAndProperties operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC33_GetListContentTypesAndProperties_ContentTypeVersion()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list succeeded.");

            #endregion Add a list

            #region Invoke GetListContentTypesAndProperties to get the old version

            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult result
                = this.listswsAdapter.GetListContentTypesAndProperties(listName, null, null, false, false);

            this.Site.Assert.IsNotNull(result, "Calling GetListContentTypesAndProperties successfully.");

            int oldVersion = result.ContentTypes.ContentType[0].Version;

            #endregion

            #region Invoke UpdateContentType

            #region Abstract Value Bind

            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(
                                                                        Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            #endregion

            #region UpdateContentType
            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResultUpdate
                = this.listswsAdapter.UpdateContentType(listName, result.ContentTypes.ContentType[0].ID, properties, addFields, null, null, bool.TrueString);

            this.Site.Assert.IsNotNull(updateContentTypeResultUpdate, "Calling GetListContentTypesAndProperties successfully.");
            #endregion

            #endregion

            #region Invoke GetListContentTypesAndProperties to get the new version

            result = this.listswsAdapter.GetListContentTypesAndProperties(listName, null, null, false, false);

            this.Site.Assert.IsNotNull(result, "Calling GetListContentTypesAndProperties successfully.");
            int newVersion = result.ContentTypes.ContentType[0].Version;

            #endregion

            #region Verify R1066

            if ((oldVersion != -1) && (newVersion != -1))
            {
                // When both the old version and the new version are got, if the 
                // new version is equal to the old version plus 1, then the following requirement 
                // can be captured.
                Site.CaptureRequirementIfAreEqual<int>(
                    oldVersion + 1,
                    newVersion,
                    1066,
                    @"[In GetListContentTypesAndPropertiesResponse]The protocol server MUST increment the value[GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.Version] by 1 each time the content type definition is edited.");
            }

            #endregion
        }

        /// <summary>
        ///  This test case is used to test the GetListContentTypesAndProperties operation with incorrect list name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC34_GetListContentTypesAndProperties_IncorrectListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create List Succeeded.");

            #endregion Add a list

            #region CreateContentType

            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "CreateContentType operation should succeed.");
            #endregion CreateContentType

            #region GetListContentTypesAndProperties
            string invalidList = DateTime.Now.TimeOfDay.ToString();
            string propertyPrefix = AdapterHelper.PrefixOws;

            try
            {
                this.listswsAdapter.GetListContentTypesAndProperties(invalidList, contextType, propertyPrefix, true, true);
                Site.Assert.Fail("GetListContentTypeAndProperties should fail, but it passed.");
            }
            catch (SoapException exp)
            {
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    1042,
                    @"[In GetListContentTypesAndProperties]If listName does not correspond to a list from either of these checks, the protocol server MUST return a SOAP fault with error code 0x82000006.");
            }

            try
            {
                this.listswsAdapter.GetList(invalidList);

                Site.Assert.Fail("Get list should fail with invalidList but it passed.");
            }
            catch (SoapException exp)
            {
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    1043,
                    @"[In [In GetListContentTypesAndProperties][If listName does not correspond to a list from either of these checks, the protocol server MUST return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetListContentTypesAndProperties operation on a UserInfoList. The ContentTypeOrder element will not be returned in Microsoft SharePoint Foundation 2010.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC35_GetListContentTypesAndProperties_OnUserInfoList()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Try to get a User Info list
            string listName = null;
            GetListCollectionResponseGetListCollectionResult listCollection = this.listswsAdapter.GetListCollection();
            foreach (ListDefinitionCT list in listCollection.Lists)
            {
                if (list.ServerTemplate == "112")
                {
                    listName = list.Name;
                    break;
                }
            }

            Site.Assert.IsNotNull(listName, "Server does not have a user info list!");

            #endregion

            #region call GetListContentTypesAndProperties
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult result = null;

            // Call GetListContentTypesAndProperties operation to get list content types and properties.
            result = this.listswsAdapter.GetListContentTypesAndProperties(
                                                                    listName,
                                                                    null,
                                                                    null,
                                                                    true,
                                                                    true);

            bool isCaptured = result.ContentTypes.ContentTypeOrder == null;

            // If result.ContentTypes.ContentTypeOrder == null, capture R3016
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                3016,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentTypeOrder: This MUST NOT be returned if the list is a user information list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test server behavior when all input parameters are valid in GetListContentTypesAndProperties operation. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC36_GetListContentTypesAndProperties_Succeed_BestMatchAndPrefix()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Data Binding
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Get the list guid of a list
            GetListResponseGetListResult getListResult = this.listswsAdapter.GetList(listName);
            Site.Assert.IsNotNull(getListResult, "Get the list id of a specified list via its list name through calling operation of GetList");
            string listGUID = getListResult.List.ID;

            // Get contentTypeId of list.
            GetListContentTypesResponseGetListContentTypesResult result;
            result = this.listswsAdapter.GetListContentTypes(listName, string.Empty);

            Site.Assert.IsNotNull(result, "The result must be not null");
            string contentTypeId = result.ContentTypes.ContentType[0].Name;
            string propertyPrefix = Common.GetConfigurationPropertyValue("PropertyPrefix", this.Site);
            #endregion

            #region Calling GetListContentTypesAndProperties operation with listName is GUID,contentTypeId specified,contentTypeId exist on the list(bestMatch),propertyPrefix not empty, includeWebProperties is true
            bool isContentTypeVerified = false;
            bool isBestMatchVerified = false;
            bool isPrefixVerified = false;

            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult resOne = null;
            resOne = this.listswsAdapter.GetListContentTypesAndProperties(listGUID, contentTypeId, propertyPrefix, true, true);

            this.Site.Assert.IsNotNull(resOne, "Call operation GetListContentTypesAndProperties should be successful.");
            #endregion

            #region Verify result, Capture R1044 R1045 R1047 R1069
            int countContentTypes = resOne.ContentTypes.ContentType.GetLength(0);
            int countProperties = resOne.Properties.GetLength(0);
            for (int i = 0; i < countContentTypes; i++)
            {
                if (resOne.ContentTypes.ContentType[i].Name == contentTypeId)
                {
                    isContentTypeVerified = true;
                    if (resOne.ContentTypes.ContentType[i].BestMatch == TRUEONLY.TRUE)
                    {
                        isBestMatchVerified = true;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isContentTypeVerified && isBestMatchVerified,
                1044,
                @"[In GetListContentTypesAndProperties]If the contentTypeId parameter is specified, the protocol server MUST attempt to find the ""best match"" to the specified content type from among those on the list.");

            Site.CaptureRequirementIfIsTrue(
                isContentTypeVerified && isBestMatchVerified,
                1045,
                @"[In GetListContentTypesAndProperties]If the content type specified by contentTypeId exists on the list, the server MUST choose that content type as the best match.");

            Site.CaptureRequirementIfIsTrue(
                isBestMatchVerified,
                1047,
                @"[In GetListContentTypesAndProperties]If the server succeeds in finding a best match to the specified content type, the server MUST set the BestMatch attribute of the selected content type to ""TRUE"" in the GetListContentTypesAndPropertiesResponse element.");

            Site.CaptureRequirementIfIsTrue(
                isContentTypeVerified && isBestMatchVerified,
                1069,
                @"[In GetListContentTypesAndPropertiesResponse]If specified at all, then it MUST be specified on the ContentType element that represents the ""best match"" to the content type identified by the contentTypeId element in the GetListContentTypesAndPropertiesSoapIn message.");

            #endregion

            #region Calling GetListContentTypesAndProperties operation with listName is GUID,contentTypeId specified,contentTypeId exist on the list(bestMatch),propertyPrefix not empty, includeWebProperties is false
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult resTwo = null;
            resTwo = this.listswsAdapter.GetListContentTypesAndProperties(listGUID, contentTypeId, propertyPrefix, false, false);

            this.Site.Assert.IsNotNull(resTwo, "Call operation GetListContentTypesAndProperties should be successful.");
            #endregion

            #region Verify result, Capture R1049
            countContentTypes = resTwo.ContentTypes.ContentType.GetLength(0);
            countProperties = resTwo.Properties.GetLength(0);

            for (int i = 0; i < countProperties; i++)
            {
                if (resTwo.Properties[i].Key.Substring(0, propertyPrefix.Length) == propertyPrefix)
                {
                    isPrefixVerified = true;
                }
                else
                {
                    isPrefixVerified = false;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isPrefixVerified,
                1049,
                @"[In GetListContentTypesAndProperties]If the propertyPrefix parameter is not empty, the protocol server MUST return all of the properties in the list property bag whose property key begins with the propertyPrefix.");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test GetListContentTypesAndProperties with correct list name GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC37_GetListContentTypesAndProperties_Succeed_WithGuid()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create List succeeded.");

            #endregion Add a list

            #region CreateContentType

            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "CreateContentType succeeded.");
            #endregion CreateContentType

            #region GetListContentTypesAndProperties
            string propertyPrefix = AdapterHelper.PrefixOws;
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult getListContentTypesAndPropertiesResult = null;

            getListContentTypesAndPropertiesResult = this.listswsAdapter.GetListContentTypesAndProperties(listId, contextType, propertyPrefix, true, true);
            if (getListContentTypesAndPropertiesResult == null)
            {
                Site.Assert.Fail("The response of the GetListContentTypesAndProperties should be not null, but actually it is null.");
            }

            bool isR1040Verified = (from contentType in getListContentTypesAndPropertiesResult.ContentTypes.ContentType
                                    where contentType.ID == contextType
                                    select contentType).ToArray().Length == 1;

            Site.CaptureRequirementIfIsTrue(
                isR1040Verified,
                1040,
                @"[In GetListContentTypesAndProperties]If listName is a valid GUID and corresponds to the identification of a list on the site, the protocol server MUST use that list");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test GetListContentTypesAndProperties with correct list name title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC38_GetListContentTypesAndProperties_Succeed_WithListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2230, this.Site), @"Test is executed only when R2230Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType

            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contentTypeId = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contentTypeId), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region GetListContentTypesAndProperties
            string propertyPrefix = AdapterHelper.PrefixOws;
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult getListContentTypesAndPropertiesResult = null;

            getListContentTypesAndPropertiesResult = this.listswsAdapter.GetListContentTypesAndProperties(listName, contentTypeId, propertyPrefix, true, true);
            if (getListContentTypesAndPropertiesResult == null)
            {
                Site.Assert.Fail("Calling GetListContentTypesAndProperties failed");
            }

            // Get the created content type object.
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResultContentTypesContentType contentType = null;
            contentType = getListContentTypesAndPropertiesResult.ContentTypes.ContentType.First(ct => ct.ID.Equals(contentTypeId, StringComparison.OrdinalIgnoreCase));

            // If the content type object in GetListContentTypesAndPropertiesResponse is the expected content type, capture R1058, R1061, R1062, R1063, R1064 
            Site.CaptureRequirementIfIsNotNull(
                contentType,
                1058,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes: The container element for the content types.");

            Site.CaptureRequirementIfAreEqual<string>(
                displayName,
                contentType.Name,
                1061,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.Name: The name of the content type.");

            Site.CaptureRequirementIfAreEqual<string>(
                contentTypeId,
                contentType.ID,
                1062,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.ID: The content type identifier for this content type.");

            // When the content type created, the display name is used for the description.
            Site.CaptureRequirementIfAreEqual<string>(
                displayName,
                contentType.Description,
                1063,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.Description: The description of the content type.");

            Uri contentTypeScope = null;
            bool isR1064Verified = Uri.TryCreate(contentType.Scope, UriKind.Absolute, out contentTypeScope);

            Site.CaptureRequirementIfIsTrue(
                isR1064Verified,
                1064,
                @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.Scope: The fully qualified URL to the content type scope.");

            // Verify R1041
            Site.CaptureRequirement(
                1041,
                @"[In GetListContentTypesAndProperties]If listName is not a valid GUID [or does not correspond to the identification of a list on the site, ]the protocol server MUST check to see if listName corresponds to the title of a list on the site and if so, MUST use that list");

            // Verify R2161
            Site.CaptureRequirement(
                2161,
                @"[If listName]does not correspond to the identification of a list on the site, the protocol server MUST check to see if listName corresponds to the title of a list on the site and if so, MUST use that list");

            if (Common.IsRequirementEnabled(2230, this.Site))
            {
                // If there is no exception generated, the R2230 is enable,
                // Then directly capture the requirement R2230.
                Site.CaptureRequirement(
                            2230,
                            "Implementation does support this operation [GetListContentTypesAndProperties]. (SharePoint Foundation 2010 and above follow this behavior.)");
            }
            #endregion
        }

        #endregion GetListContentTypesAndProperties

        #region GetListContentTypes

        /// <summary>
        /// This test case is used to test the content type's version on the protocol server in UpdateContentType operation and GetListContentTypes operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC39_GetListContentTypes_VersionTest()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create List succeeded.");

            #endregion Add a list

            #region Invoke GetListContentTypes to get the old version

            GetListContentTypesResponseGetListContentTypesResult result = this.listswsAdapter.GetListContentTypes(listName, null);

            this.Site.Assert.IsNotNull(result, "GetListContentTypes Succeeded.");

            int oldVersion = result.ContentTypes.ContentType[0].Version;

            #endregion

            #region Abstract Value Bind

            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            #endregion

            #region Invoke UpdateContentType
            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResultUpdate
                = this.listswsAdapter.UpdateContentType(listName, result.ContentTypes.ContentType[0].ID, properties, addFields, null, null, bool.TrueString);

            this.Site.Assert.IsNotNull(updateContentTypeResultUpdate, "UpdateContentType successfully.");
            #endregion

            #region Invoke GetListContentTypes to get the new version

            result = this.listswsAdapter.GetListContentTypes(listName, string.Empty);

            this.Site.Assert.IsNotNull(result, "GetListContentTypes Succeeded.");

            int newVersion = result.ContentTypes.ContentType[0].Version;

            #endregion

            #region Verify 619

            if ((oldVersion != -1) && (newVersion != -1))
            {
                // When both the old version and the new version are obtained, the following requirement                
                // can be captured if the new version is equal to the old version plus 1.
                Site.CaptureRequirementIfAreEqual<int>(
                    oldVersion + 1,
                    newVersion,
                    619,
                    @"[GetListContentTypesResult.ContentTypes.ContentType.Version] The protocol server MUST increment the value by 1 each time the content type definition is edited.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to test a BestMatch content type is returned in the response in GetListContentTypes operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC40_GetListContentTypes_BestMatch()
        {
            #region Add a list

            string listName = TestSuiteHelper.CreateList();

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listName);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create List succeeded.");
            #endregion Add a list
            GetListContentTypesResponseGetListContentTypesResult listContentTypes = this.listswsAdapter.GetListContentTypes(listName, null);

            this.Site.Assert.IsNotNull(listContentTypes, "GetListContentType succeeded.");
            string contentTypeId = null;
            for (int i = 0; i < listContentTypes.ContentTypes.ContentType.Length; i++)
            {
                if (listContentTypes.ContentTypes.ContentType[i].Name.Equals("Document"))
                {
                    contentTypeId = listContentTypes.ContentTypes.ContentType[i].ID;
                    break;
                }
            }

            // GetListContentTypes using the list's title and specifying Document content type is "best match".
            listContentTypes = this.listswsAdapter.GetListContentTypes(listName, contentTypeId);

            Site.Assert.IsNotNull(listContentTypes, @"[In GetListContentTypes operation] If the specified listName a list name, the protocol server MUST return a GetListResponse element.");

            // BestMatch is specified as TRUE at all, since it is TRUEONLY type.
            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in listContentTypes.ContentTypes.ContentType)
            {
                if (ct.Name == "Document")
                {
                    Site.Assert.IsTrue(ct.BestMatchSpecified, "The content type identified by the contentTypeId element in the GetListContentTypesSoapIn message should be specified as best match.");
                }
                else
                {
                    Site.Assert.IsFalse(ct.BestMatchSpecified, "The content type not identified by the contentTypeId element in the GetListContentTypesSoapIn message should not be specified as best match.");
                }
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1844
            // Since the asserts above are successful,MS-LISTSWS_R1844 can be captured directly.
            Site.CaptureRequirement(
                1844,
                @"[GetListContentTypesResult.ContentTypes.ContentType.BestMatch]If specified at all, then MUST be specified on the ContentType element that represents the ""best match"" to the content type identified by the contentTypeId element in the GetListContentTypesSoapIn message.");
        }

        /// <summary>
        /// This test case is used to test the input contentTypeId is an identifier of a content type in GetListContentTypes operation whether the BestMatch attribute of the content type is specified in the response.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC41_GetListContentTypes_GetBestMatch()
        {
            #region Add a list
            string listName = TestSuiteHelper.CreateList();

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listName);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list succeeded.");

            #endregion Add a list

            #region Get all content types in the list
            GetListContentTypesResponseGetListContentTypesResult getListContentTypesResult = this.listswsAdapter.GetListContentTypes(listName, string.Empty);

            this.Site.Assert.IsNotNull(getListContentTypesResult, "GetListContentTypes succeeded.");
            #endregion

            #region Use one of the content type ID to get the BestMatch

            string contentTypeId = getListContentTypesResult.ContentTypes.ContentType[0].ID;

            getListContentTypesResult = this.listswsAdapter.GetListContentTypes(listName, contentTypeId);

            this.Site.Assert.IsNotNull(getListContentTypesResult, "GetListContentType succeeded.");
            #endregion

            #region Verify R613

            // If one of the returned content type has an ID equal to the input value 
            // and the BestMatchSpecified is True, then the following requirement can be captured.
            bool isR613Verified = false;
            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in getListContentTypesResult.ContentTypes.ContentType)
            {
                if (ct.BestMatchSpecified && (ct.ID == contentTypeId))
                {
                    isR613Verified = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isR613Verified,
                613,
                @"[In GetListContentTypes operation] If the server succeeds in finding a best match to the specified content type, the server MUST set the BestMatch attribute of the selected content type to ""TRUE"" in the GetListContentTypesResponse element.");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test GetListContentTypes with incorrect list name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC42_GetListContentTypes_IncorrectListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2762, this.Site), @"Test is executed only when R2762Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create List succeeded.");

            #endregion Add a list

            #region CreateContentType

            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "CreateContentType Succeeded.");
            #endregion CreateContentType

            #region GetListContentTypes
            string invalidList = DateTime.Now.TimeOfDay.ToString();
            bool caughtSoapException = false;
            try
            {
                this.listswsAdapter.GetListContentTypes(invalidList, string.Empty);
                Site.Assert.Fail("GetListContentType should fail");
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    2762,
                    @"[In GetListContentTypes operation] Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 and above follow this behavior.)");

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    1825,
                    @"[In GetListContentTypes operation] [If listName does not correspond to a list from either of these checks, the protocol server should<51> return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            }
            finally
            {
                Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetListContentTypes' with  incorrect list name.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetListContentTypes operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC43_GetListContentTypes_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2460, this.Site), @"Test is executed only when R2460Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                // Call GetListContentTypes method to get retrieve all content types from a list.
                this.listswsAdapter.GetListContentTypes(invalidListName, string.Empty);
            }
            catch (SoapException soapException)
            {
                bool isVerifyR2460 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2460 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2460 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2460,
                    2460,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<51> Section 3.1.4.19.2.1: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to test none BestMatch content type is returned in the response in GetListContentTypes operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC44_GetListContentTypes_NoBestMatch()
        {
            #region Add a list
            string listName = TestSuiteHelper.CreateList();

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listName);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list succeeded.");
            #endregion Add a list

            #region GetListContentTypes using the list's GUID,specifing a "best match" field with an invalid content type id that the server cannot find.
            string contentTypeId = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            GetListContentTypesResponseGetListContentTypesResult listContentTypes = this.listswsAdapter.GetListContentTypes(listName, contentTypeId);

            Site.Assert.IsNotNull(listContentTypes, @"[In GetListContentTypes operation] If the specified listName a valid GUID, the protocol server MUST return a GetListResponse element.");
            int bestMatchCount = 0;
            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in listContentTypes.ContentTypes.ContentType)
            {
                if (ct.BestMatchSpecified == true)
                {
                    bestMatchCount++;
                }
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1828
            Site.CaptureRequirementIfAreEqual<int>(0, bestMatchCount, 1828, @"[GetListContentTypes]If the server is unable to find a ""best match,"" the server MUST NOT set the BestMatch attribute to ""TRUE"" for any of the content types returned in the GetListContentTypesResponse element.");
            #endregion
        }

        /// <summary>
        /// This test case is used to validate that the server does not return ContentTypeOrder element when GetListContentTypes operation functions on a UserInfoList.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC45_GetListContentTypes_OnUserInfoList()
        {
            #region Try to get a User Info list
            string listName = null;
            GetListCollectionResponseGetListCollectionResult listCollection = this.listswsAdapter.GetListCollection();
            foreach (ListDefinitionCT list in listCollection.Lists)
            {
                if (list.ServerTemplate == "112")
                {
                    listName = list.Name;
                    break;
                }
            }

            Site.Assert.IsNotNull(listName, "Server does not have a user info list!");

            #endregion

            // Call GetListContentTypes operation to get list content types.
            GetListContentTypesResponseGetListContentTypesResult contentTypes = this.listswsAdapter.GetListContentTypes(listName, null);

            bool isR1834Captured = contentTypes.ContentTypes.ContentTypeOrder == null;

            // If contentTypes.ContentTypes.ContentTypeOrder == null, capture R1834.
            Site.CaptureRequirementIfIsTrue(
                isR1834Captured,
                1834,
                @"[GetListContentTypesResult.ContentTypes.ContentTypeOrder] This MUST NOT be returned if the list is a user information list.");
        }

        /// <summary>
        /// This test case is used to test GetListContentTypes about finding best match.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC46_GetListContentTypes_Succeed_BestMatch()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contentTypeId = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contentTypeId), "CreateContentType successfully.");
            #endregion CreateContentType

            #region GetListContentTypes
            GetListContentTypesResponseGetListContentTypesResult listContentTypes = this.listswsAdapter.GetListContentTypes(listId, contentTypeId);
            this.Site.Assert.IsNotNull(listContentTypes, "Calling GetListContentTypes successfully.");
            bool isListExist = listContentTypes.ContentTypes.ContentType.Any(contentType => contentType.ID.Equals(contentTypeId, StringComparison.OrdinalIgnoreCase));

            if (!isListExist)
            {
                Site.Assert.Fail("There is no list which matches the list name{0} in the GetListContentTypes response", listName);
            }

            bool isBestMatchExist = false;

            // BestMatch is specified as TRUE 
            foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in listContentTypes.ContentTypes.ContentType)
            {
                if (ct.BestMatchSpecified == true)
                {
                    isBestMatchExist = true;
                    break;
                }
            }

            if (!isBestMatchExist)
            {
                Site.Assert.Fail("Could not find the Best Match.");
            }

            // If BestMatch is specified as TRUE, then capture R612.
            Site.CaptureRequirement(
                612,
                @"[In GetListContentTypes operation] If the contentTypeId parameter is specified, the protocol server MUST attempt to find the ""best match"" to the specified content type from among those on the list.");

            // If BestMatch is specified as TRUE, then capture R1826.
            Site.CaptureRequirement(
                1826,
                @"[GetListContentTypes]If the content type specified by contentTypeId exists on the list, the server MUST choose that content type as the best match.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetListContentTypes with correct list name GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC47_GetListContentTypes_Succeed_WithGuid()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contentTypeId = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contentTypeId), "CreateContentType successfully.");
            #endregion CreateContentType

            #region GetListContentTypes
            GetListContentTypesResponseGetListContentTypesResult listContentTypes = this.listswsAdapter.GetListContentTypes(listId, string.Empty);
            this.Site.Assert.IsNotNull(listContentTypes, "Calling GetListContentTypes successfully.");

            // Verify requirement R6101.
            // If getting content type is successful, it means implementation does support this GetListContentTypes method. R6101 can be captured.
            if (Common.IsRequirementEnabled(6101, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    listContentTypes,
                    6101,
                    @"Implementation does support this method[GetListContentTypes]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            bool isR608Verified = listContentTypes.ContentTypes.ContentType.Any(contentType => contentType.ID.Equals(contentTypeId, StringComparison.OrdinalIgnoreCase));

            // If the create content type can be found in GetListContentTypes response, that means the list is the expected list, Then capture R608.
            Site.CaptureRequirementIfIsTrue(
                isR608Verified,
                608,
                @"[In GetListContentTypes operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, the protocol server MUST use that list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test GetListContentTypes with correct list name title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC48_GetListContentTypes_Succeed_WithListName()
        {
            #region Add a list

            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();

            // This field has been added into current list.
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            List<string> fields = new List<string> { fieldName };

            string contentTypeId = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contentTypeId), "CreateContentType successfully.");
            #endregion CreateContentType

            #region GetListContentTypes
            GetListContentTypesResponseGetListContentTypesResult listContentTypes = this.listswsAdapter.GetListContentTypes(listName, string.Empty);
            this.Site.Assert.IsNotNull(listContentTypes, "Calling GetListContentTypes successfully.");

            if (Common.IsRequirementEnabled(18321, this.Site))
            {
                string[] contentTypeOrderArray = listContentTypes.ContentTypes.ContentTypeOrder.Split(',');

                bool isR18321Verified = true;
                foreach (string contentTypeOrder in contentTypeOrderArray)
                {
                    isR18321Verified = isR18321Verified && listContentTypes.ContentTypes.ContentType.Any(ct => ct.ID == contentTypeOrder);
                }

                // If each split string is a valid content type ID, capture R18321.
                Site.CaptureRequirementIfIsTrue(
                    isR18321Verified,
                    18321,
                    @"[GetListContentTypesResponse]Implementation does return a string of content type identifiers delimited by commas.(SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Get the created content type object.
            GetListContentTypesResponseGetListContentTypesResultContentTypesContentType contentType = null;
            contentType = listContentTypes.ContentTypes.ContentType.First(ct => ct.ID.Equals(contentTypeId, StringComparison.OrdinalIgnoreCase));

            // If the content type object in GetListContentTypesResponse is the expected content type, capture R1831, R1838, R1839, R1840, R1841 
            Site.CaptureRequirementIfIsNotNull(
                contentType,
                1831,
                @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes: The container element for the content types.");

            Site.CaptureRequirementIfAreEqual<string>(
                displayName,
                contentType.Name,
                1838,
                @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes.ContentType.Name: The name of the content type.");

            Site.CaptureRequirementIfAreEqual<string>(
                contentTypeId,
                contentType.ID,
                1839,
                @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes.ContentType.ID: The content type identifier for this content type.");

            // When the content type created, the display name is used for the description.
            Site.CaptureRequirementIfAreEqual<string>(
                displayName,
                contentType.Description,
                1840,
                @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes.ContentType.Description: The description of the content type.");

            Uri contentTypeScope = null;
            bool isR1841Verified = Uri.TryCreate(contentType.Scope, UriKind.Absolute, out contentTypeScope);

            Site.CaptureRequirementIfIsTrue(
                isR1841Verified,
                1841,
                @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes.ContentType.Scope: The fully qualified URL to the content type scope.");

            // If the create content type can be found in GetListContentTypes response, that means the list is the expected list, Then capture R609 R610.
            bool isR609Verified = listContentTypes.ContentTypes.ContentType.Any(ct => ct.ID.Equals(contentTypeId, StringComparison.OrdinalIgnoreCase));

            Site.CaptureRequirementIfIsTrue(
                isR609Verified,
                609,
                @"[In GetListContentTypes operation] If the specified listName is not a valid GUID, the protocol server MUST check to see if listName corresponds to the title of a list on the site and if so, MUST use that list.");

            Site.CaptureRequirementIfIsTrue(
                isR609Verified,
                610,
                @"[In GetListContentTypes operation] If the specified listName does not correspond to the identification of a list on the site, the protocol server MUST check to see if listName corresponds to the title of a list on the site and if so, MUST use that list.");

            #endregion
        }

        #endregion

        #region UpdateContentTypeXMLDocument

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with incorrect contentTypeID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC49_UpdateContentTypeXMLDocument_IncorrectContentTypeId()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");

            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument
            // Construct the newDocument parameter for the invocation below.
            string qualifiedNameOne = TestSuiteHelper.GetUniqueFieldName();
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(qualifiedNameOne, namespaceURIOne, innerXmlOne);
            string invalidContextType = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            try
            {
                this.listswsAdapter.UpdateContentTypeXmlDocument(listId, invalidContextType, newDocument);
                Site.Assert.Fail("Calling UpdateContentTypeXmlDocument should fail");
            }
            catch (SoapException ex)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(ex);

                Site.CaptureRequirementIfIsNotNull(
                    ex,
                    867,
                    @"[In UpdateContentTypeXmlDocument operation] If the content type specified by the contentTypeId is not found, the protocol server MUST return a SOAP fault.");

                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    2027,
                    @"[In UpdateContentTypeXmlDocument operation] [If the content type specified by the contentTypeId is not found, the protocol server MUST return a SOAP fault.] There is no error code for this fault.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with incorrect listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC50_UpdateContentTypeXMLDocument_IncorrectListName()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument
            // Construct the newDocument parameter for the invocation below.
            string qualifiedNameOne = TestSuiteHelper.GetUniqueFieldName();
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(qualifiedNameOne, namespaceURIOne, innerXmlOne);

            string invalidList = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool issoapFaultGenerated = false;
            try
            {
                this.listswsAdapter.UpdateContentTypeXmlDocument(invalidList, contextType, newDocument);
                Site.Assert.Fail("Calling UpdateContentType should fail");
            }
            catch (SoapException exp)
            {
                issoapFaultGenerated = true;
                if (Common.IsRequirementEnabled(2768, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<string>(
                       "0x82000006",
                       exp.Detail.LastChild.InnerText,
                       2768,
                       @"[In UpdateContentTypeXmlDocument operation] Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");

                    Site.CaptureRequirementIfAreEqual<string>(
                       "0x82000006",
                       exp.Detail.LastChild.InnerText,
                       2026,
                       @"[In UpdateContentTypeXmlDocument operation] [If listName does not correspond to a list from either of these checks, the protocol server should<77> return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
                }
            }
            finally
            {
                this.Site.CaptureRequirementIfIsTrue(
                          issoapFaultGenerated,
                          878,
                          @"[In UpdateContentTypeXmlDocument operation] [In UpdateContentTypeXmlDocumentResponse element] [In UpdateContentTypeXmlDocumentResult element] If the operation [UpdateContentTypeXmlDocument] fails, a SOAP fault MUST be returned .");
            }

            #endregion UpdateContentTypeXMLDocument
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with correct listName GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC51_UpdateContentTypeXMLDocument_WithListGuid_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument
            // Construct the newDocument parameter for the invocation below.
            string qualifiedNameOne = Common.GetConfigurationPropertyValue("ContentTypeXmlQualifiedName", this.Site);
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);

            XmlDocument doc = new XmlDocument();
            XmlNode newDocument = null;

            newDocument = doc.CreateNode(XmlNodeType.DocumentFragment, qualifiedNameOne, namespaceURIOne);
            newDocument.InnerXml = innerXmlOne;

            System.Xml.XPath.IXPathNavigable updateContentTypeXmlDocumentResult
                = this.listswsAdapter.UpdateContentTypeXmlDocument(listId, contextType, newDocument);

            // If the UpdateContentTypeXmlDocument operation succeeds, then the server will return a successful response.
            // So if the response isn't null, then capture R863.
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                863,
                @"[In UpdateContentTypeXmlDocument operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            #endregion UpdateContentTypeXMLDocument
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with correct listName title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC52_UpdateContentTypeXMLDocument_WithListTitle_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);

            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypeXMLDocument
            // Construct the newDocument parameter for the invocation below.
            string qualifiedNameOne = Common.GetConfigurationPropertyValue("ContentTypeXmlQualifiedName", this.Site);
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);

            XmlDocument doc = new XmlDocument();
            XmlNode newDocument = null;

            newDocument = doc.CreateNode(XmlNodeType.DocumentFragment, qualifiedNameOne, namespaceURIOne);
            newDocument.InnerXml = innerXmlOne;

            System.Xml.XPath.IXPathNavigable updateContentTypeXmlDocumentResult
                = this.listswsAdapter.UpdateContentTypeXmlDocument(listName, contextType, newDocument);

            // If the UpdateContentTypeXmlDocument updates the list whose list title corresponds to the specified listName successfully, the server will return a successful response.
            // So if the response of the operation isn't null, then capture R864 and R865.
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                864,
                @"[In UpdateContentTypeXmlDocument operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                865,
                @"[In UpdateContentTypeXmlDocument operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                877,
                @"[In UpdateContentTypeXmlDocument operation] [In UpdateContentTypeXmlDocumentResponse element] [In UpdateContentTypeXmlDocumentResult element] If the operation succeeds, an UpdateContentTypeXmlDocumentResult MUST be returned.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                869,
                @"[In UpdateContentTypeXmlDocument operation] Otherwise [If there is no XML document with the same namespace URI exist], the new XML document MUST be added to the content type XML document.");

            // Update existing XML document
            // Used the same namespace URI
            string qualifiedNameNew = TestSuiteHelper.GetUniqueFieldName();
            newDocument = doc.CreateNode(XmlNodeType.DocumentFragment, qualifiedNameNew, namespaceURIOne);
            newDocument.InnerXml = innerXmlOne;

            updateContentTypeXmlDocumentResult
                = this.listswsAdapter.UpdateContentTypeXmlDocument(listName, contextType, newDocument);

            // If the UpdateContentTypeXmlDocument operation succeeds, the server will return a successful response.
            // So if the response isn't null, then capture R868.
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                868,
                @"[In UpdateContentTypeXmlDocument operation] If an XML document with the same namespace URI [RFC2396] already exists, it MUST be replaced by the new XML document.");

            #endregion UpdateContentTypeXMLDocument
        }

        /// <summary>
        /// This test case is used to verify UpdateContentTypeXmlDocument operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC53_UpdateContentTypeXmlDocument_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2502, this.Site), @"Test is executed only when R2502Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;

            // Read value from configuration file.
            string namespaceUri1 = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            string elementName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Initialize a newDocument.
            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, namespaceUri1, innerXml);
            try
            {
                // Call UpdateContentTypeXmlDocument method to update XML Document properties of the content type collection on a list.
                this.listswsAdapter.UpdateContentTypeXmlDocument(invalidListName, string.Empty, newDocument);
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2502
                bool isVerifyR2502 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2502 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2502 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2502,
                    2502,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<77> Section 3.1.4.29: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to verify that an xml document is successfully added into content type in UpdateContentTypeXmlDocument.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC54_UpdateContentTypeXmlDocument_Success()
        {
            // Create a new list
            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC53_UpdateContentTypeXmlDocument_Success, create a new list successfully.");

            // Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");

            // Call GetListContentType operation to get the new content type.
            GetListContentTypeResponseGetListContentTypeResult getContentTypeResult = null;
            getContentTypeResult = this.listswsAdapter.GetListContentType(listName, contentTypeId);
            int originalXmlDocNum = getContentTypeResult.ContentType.XmlDocuments.Length;

            // Initialize a newDocument with the existing namespace URI.
            string namespaceUri1 = getContentTypeResult.ContentType.XmlDocuments[0].NamespaceURI;
            string innerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            string elementName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            XmlNode newDocument = TestSuiteHelper.CreateNewDocument(elementName, namespaceUri1, innerXml);

            // Call UpdateContentTypeXmlDocument to update the content type xml document.
            System.Xml.XPath.IXPathNavigable updateContentTypeXmlDocumentResult = null;
            updateContentTypeXmlDocumentResult = this.listswsAdapter.UpdateContentTypeXmlDocument(listName, contentTypeId, newDocument);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R8581.
            // If updating content type XML document is successful, it means implementation does support this UpdateContentTypeXmlDocument method. R8581 can be captured.
            if (Common.IsRequirementEnabled(8581, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    updateContentTypeXmlDocumentResult,
                    8581,
                    @"Implementation does support this method[UpdateContentTypeXmlDocument]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // Call GetListContentType operation again to get the updated content type.
            getContentTypeResult = this.listswsAdapter.GetListContentType(listName, contentTypeId);

            // Check if the total number of XML document are same in the 2 GetListContentType responses.
            bool isSameXmlDocNum = originalXmlDocNum == getContentTypeResult.ContentType.XmlDocuments.Length;

            // Check if the new XML document can be found in the 2nd GetListContentType response.
            XmlDocumentDefinition xmlDoc = getContentTypeResult.ContentType.XmlDocuments[0];
            bool isNewDocExist = xmlDoc.Any.Any(xmlNode => xmlNode.InnerXml == newDocument.InnerXml);

            // If the total number of XML document are same in the 2 GetListContentType responses, and the new XML document can be 
            // found in the 2nd GetListContentType response, that means the original XML document is replaced, then capture R847.
            Site.CaptureRequirementIfIsTrue(
                isSameXmlDocNum && isNewDocExist,
                847,
                @"[In UpdateContentTypesXmlDocument operation] If an XML document with the same namespace URI [RFC2396] already exists, it MUST be replaced by the new XML Document.");
        }

        #endregion

        #region UpdateContentType

        /// <summary>
        /// This test case is used to verify the negative status of adding a new field but success of adding, updating and deleting of existing fields in UpdateContentType operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC55_UpdateContentType_AddFieldsError()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC54_UpdateContentType_AddFieldsError, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldNameB = TestSuiteHelper.GetUniqueFieldName();
            string fieldNameC = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA, fieldNameC });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            // Call GetListContentType operation to get the content type.
            this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);

            // Construct one invalid added fields
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(fieldNameB);

            // Construct one valid update field
            AddOrUpdateFieldsDefinition updateFields = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameA);

            // Construct one valid delete field
            DeleteFieldsDefinition deleteFields = TestSuiteHelper.CreateDeleteContentTypeFields(fieldNameC);

            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();

            // Call UpdateContentType operation to update the content type. updateFields and deleteFields should success, addFields should fail.
            UpdateContentTypeResponseUpdateContentTypeResult updateResult = this.listswsAdapter.UpdateContentType(
                                                   listName,
                                                   contentTypeId,
                                                   updateProperties,
                                                   addFields,
                                                   updateFields,
                                                   deleteFields,
                                                   Convert.ToString(true));

            bool isCaptured = updateResult.Results.Method[1].ErrorCode == "0x82000007";

            // Capture R807 if the error code is "0x82000007".
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                807,
                @"[In UpdateContentType operation] [The protocol server adds the appropriate fields to the content type.] If the fields to be added do not exist in the specified list, the protocol server MUST return a SOAP faultResponse with error code 0x82000007. ");
        }

        /// <summary>
        /// This test case is used to test successfully adding and deleting fields in UpdateContentType operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC56_UpdateContentType_Combination()
        {
            string lisid = TestSuiteHelper.CreateList();

            // add a field into the list.
            string addedFieldName = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(lisid, new List<string> { addedFieldName }, new List<string> { "Counter" }, new List<string> { null });
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();

            // Create a content type
            string existringFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contenttypeId = TestSuiteHelper.CreateContentType(lisid, displayName, new List<string> { existringFieldName });

            // Add a field into the content type
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(addedFieldName);
            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();
            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResult = null;
            updateContentTypeResult = this.listswsAdapter.UpdateContentType(
                                            lisid,
                                            contenttypeId,
                                            properties,
                                            addFields,
                                            null,
                                            null,
                                            bool.TrueString);

            if (null == updateContentTypeResult || null == updateContentTypeResult.Results || null == updateContentTypeResult.Results.Method)
            {
                this.Site.Assert.Fail("Could not get the valid response from UpdateContentType operation");
            }

            bool isAddFieldsErrorCodeCorrect = "0x00000000".Equals(updateContentTypeResult.Results.Method[0].ErrorCode, StringComparison.OrdinalIgnoreCase);

            // Search the method items which are executed successfully
            var fieldRefsOfSuccessfulMethodItem = from methoditem in updateContentTypeResult.Results.Method
                                                  where "0x00000000".Equals(methoditem.ErrorCode, StringComparison.OrdinalIgnoreCase)
                                                        && methoditem.FieldRef != null
                                                  select methoditem.FieldRef;

            // Search the matched method item in  items which are executed successfully
            var fieldRefsOfMatchFieldName = from fieldItem in fieldRefsOfSuccessfulMethodItem
                                            where addedFieldName.Equals(fieldItem.Name, StringComparison.OrdinalIgnoreCase)
                                            select fieldItem;

            this.Site.CaptureRequirementIfAreEqual(
                1,
                fieldRefsOfMatchFieldName.Count(),
                832,
                @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In FieldRef element] If the ErrorCode is ""0x00000000"" for an add operation, FieldRef MUST contain data for the reference to the field.");

            GetListContentTypeResponseGetListContentTypeResult getContentTypeResult = null;
            getContentTypeResult = this.listswsAdapter.GetListContentType(lisid, contenttypeId);

            if (null == getContentTypeResult || null == getContentTypeResult.ContentType || null == getContentTypeResult.ContentType.Fields
                || null == getContentTypeResult.ContentType.Fields.Items)
            {
                this.Site.Assert.Fail("Could not get the valid response from GetListContentType operation");
            }

            // Verify whether the field is successfully added into the content type.
            var matchedFielditems = from FieldDefinitionCTRemove fielditem in getContentTypeResult.ContentType.Fields.Items
                                    where addedFieldName.Equals(fielditem.Name, StringComparison.OrdinalIgnoreCase)
                                    select fielditem;

            this.Site.CaptureRequirementIfAreEqual(
                1,
                matchedFielditems.Count(),
                2273,
                @"[In UpdateContentType operation] The protocol server adds the appropriate fields to the content type.");

            // Delete an existing field on a contentType
            DeleteFieldsDefinition deleteFields = TestSuiteHelper.CreateDeleteContentTypeFields(addedFieldName);
            updateContentTypeResult = this.listswsAdapter.UpdateContentType(
                                             lisid,
                                             contenttypeId,
                                             properties,
                                             null,
                                             null,
                                             deleteFields,
                                             bool.TrueString);

            if (null == updateContentTypeResult || null == updateContentTypeResult.Results || null == updateContentTypeResult.Results.Method)
            {
                this.Site.Assert.Fail("Could not get the valid response from UpdateContentType operation");
            }

            bool isDeleteFieldsErrorCodeCorrect = "0x00000000".Equals(updateContentTypeResult.Results.Method[0].ErrorCode, StringComparison.OrdinalIgnoreCase);

            // Verify requirement #829
            this.Site.CaptureRequirementIfIsTrue(
                isAddFieldsErrorCodeCorrect && isDeleteFieldsErrorCodeCorrect,
                829,
                @"[In UpdateContentType operation] For an add or delete operation, if the field to be added or deleted is found on the list or context site, the value of ErrorCode MUST be ""0x00000000"".");

            // Verify whether the field is deleted successfully
            var executedSuccessfullyMethodItem = from methoditem in updateContentTypeResult.Results.Method
                                                 where "0x00000000".Equals(methoditem.ErrorCode, StringComparison.OrdinalIgnoreCase)
                                                 select methoditem.FieldRef;

            this.Site.Assert.AreEqual(
                                    1,
                                    executedSuccessfullyMethodItem.Count(),
                                    "The DeletedFields method was not executed successfully.");

            getContentTypeResult = null;
            getContentTypeResult = this.listswsAdapter.GetListContentType(lisid, contenttypeId);
            if (null == getContentTypeResult || null == getContentTypeResult.ContentType || null == getContentTypeResult.ContentType.Fields
               || null == getContentTypeResult.ContentType.Fields.Items)
            {
                this.Site.Assert.Fail("Could not get the valid response from GetListContentType operation");
            }

            // Verify whether the field is successfully added into the content type 
            var leftields = from FieldDefinitionCTRemove fielditem in getContentTypeResult.ContentType.Fields.Items
                            where addedFieldName.Equals(fielditem.Name, StringComparison.OrdinalIgnoreCase)
                            select fielditem;

            // If the deleted field does not exist in response of GetContenttype operation, R804 is captured.
            this.Site.CaptureRequirementIfAreEqual(
                                            0,
                                            leftields.Count(),
                                            804,
                                            @"[In UpdateContentType operation] If the specified content type is found, the protocol server deletes the appropriate fields from the content type.");
        }

        /// <summary>
        /// This test case is used to verify the negative status of deleting a field but success of adding, updating and deleting of other fields in UpdateContentType operation.  
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC57_UpdateContentType_DeleteFieldsError()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC56_UpdateContentType_DeleteFieldsError, create a new list successfully.");
            #endregion

            #region Create an content type on the list

            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldNameB = TestSuiteHelper.GetUniqueFieldName();
            string fieldNameC = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            // Call GetListContentType operation to get the content type.
            this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);

            // Construct two delete fields, one is valid and the other one is invalid
            DeleteFieldsDefinition deleteFields = TestSuiteHelper.CreateDeleteContentTypeFields(fieldNameB);

            // Construct one valid update field
            AddOrUpdateFieldsDefinition updateField = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameA);

            // Construct one valid add field
            AddOrUpdateFieldsDefinition addField = TestSuiteHelper.CreateAddContentTypeFields(fieldNameC);

            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();

            // Call UpdateContentType operation to update the content type.
            UpdateContentTypeResponseUpdateContentTypeResult updateResult = this.listswsAdapter.UpdateContentType(
                                                    listName,
                                                    contentTypeId,
                                                    updateProperties,
                                                    addField,
                                                    updateField,
                                                    deleteFields,
                                                    Convert.ToString(true));

            Site.Assert.IsNotNull(updateResult, "The result of UpdateContentType operation should not be NULL");

            bool isCaptured = updateResult.Results.Method[0].ErrorCode == "0x82000007";

            // Capture R805 if the error code is "0x82000007".
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                805,
                @"[In UpdateContentType operation][If the specified content type is found, the protocol server deletes the appropriate fields from the content type. ]If the fields to be deleted do not exist in the specified list, the protocol server MUST return a Response with error code 0x82000007. ");

            // Call GetListContentType operation to get the content type.
            GetListContentTypeResponseGetListContentTypeResult getResult = this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);

            Site.Assert.IsNotNull(getResult, "The result of GetListContentType operation should not be NULL");
        }

        /// <summary>
        /// This test case is used to validate the UpdateContentType operation when the field name is existent.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC58_UpdateContentType_FieldExists()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();

            // Create a new list.
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC57_UpdateContentType_FieldExists, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            // Create a content type.
            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            #region UpdateContentType Operation when the update field name exists

            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();
            updateProperties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition updateField = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameA);
            UpdateContentTypeResponseUpdateContentTypeResult result = null;

            // Call UpdateContentType operation to update the content type.
            result = this.listswsAdapter.UpdateContentType(
                                            listName,
                                            contentTypeId,
                                            updateProperties,
                                            null,
                                            updateField,
                                            null,
                                            Convert.ToString(true));

            Site.Assert.IsNotNull(result, "The result of UpdateContentType operation should not be NULL");

            bool isCaptured = result.Results.Method[0].ErrorCode == "0x00000000";

            // Capture R830 if the error code is "0x00000000".
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                830,
                @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In ErrorCode element] For an update operation, if a reference to the field to be updated is found on the content type, the value MUST be ""0x00000000"".");

            // Construct an FieldDefinition, to verify the value of the Field.
            FieldDefinition dieldDef = result.Results.Method[0].Field;

            // Capture R833 if the Field is contains data.
            Site.CaptureRequirementIfIsNotNull(
                dieldDef,
                833,
                @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In Field element] If the ErrorCode is ""0x00000000"" for an update operation, Field MUST contain data for the Field.");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test that the error code will be non-zero when using UpdateContentType method to delete a field that does not exist in the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC59_UpdateContentType_FieldNotExists()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();

            // Create a new list.
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC58_UpdateContentType_FieldNotExists, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            #region UpdateContentType when the update field name does not exist

            string fieldNameB = TestSuiteHelper.GetUniqueFieldName();
            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();
            updateProperties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition updateField = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameB);
            UpdateContentTypeResponseUpdateContentTypeResult result = null;

            // Update created content type in the previous step without exist field name
            result = this.listswsAdapter.UpdateContentType(
                                            listName,
                                            contentTypeId,
                                            updateProperties,
                                            null,
                                            updateField,
                                            null,
                                            Convert.ToString(true));

            Site.Assert.IsNotNull(result, "The result of UpdateContentType operation should not be NULL");

            bool isCaptured = result.Results.Method[0].ErrorCode != "0x00000000";

            // Capture R831 if the error code is not "0x00000000".
            Site.CaptureRequirementIfIsTrue(
            isCaptured,
            831,
            @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In ErrorCode element] Otherwise [in any scenarios except the successful scenarios], the value of ErrorCode MUST be nonzero.");

            // Capture R834 if the error text is not null.
            Site.CaptureRequirementIfIsNotNull(
                result.Results.Method[0].ErrorText,
                834,
                @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In ErrorText element] If the ErrorCode is nonzero, ErrorText MUST contain an error string.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with incorrect contentTypeId.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC60_UpdateContentType_IncorrectContentTypeId()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list should succeed.");

            #endregion Add a list

            #region Invoke GetListContentTypes to get the old version

            GetListContentTypesResponseGetListContentTypesResult result = this.listswsAdapter.GetListContentTypes(listId, null);

            this.Site.Assert.IsNotNull(result, "GetListContentTypes operation should succeed.");
            #endregion

            #region Abstract Value Bind
            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            #endregion

            #region UpdateContentType

            // incorrect contentTypeID
            string invalidContentTypeID = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            try
            {
                this.listswsAdapter.UpdateContentType(listId, invalidContentTypeID, properties, addFields, null, null, bool.TrueString);
                Site.Assert.Fail("UpdateContentType should fail when the parent content type ID is invalid, but it passed.");
            }
            catch (SoapException ex)
            {
                Site.CaptureRequirementIfIsNotNull(
                    ex,
                    801,
                    @"[In UpdateContentType operation] If the contentTypeId is not a valid content type identifier, the protocol server MUST return a SOAP fault.");
            }

            // contentTypeId is not the content type identifier of a content type on the list
            invalidContentTypeID = Guid.NewGuid().ToString("N");

            try
            {
                this.listswsAdapter.UpdateContentType(listId, invalidContentTypeID, properties, addFields, null, null, bool.TrueString);
                Site.Assert.Fail("Calling UpdateContentType should fail");
            }
            catch (SoapException ex)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(ex);

                Site.CaptureRequirementIfIsNotNull(
                    ex,
                    802,
                    @"[In UpdateContentType operation] If the contentTypeId is not the content type identifier of a content type on the list, the protocol server MUST return a SOAP fault.");

                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    1976,
                    @"[In UpdateContentType operation] [If the contentTypeId is not the content type identifier of a content type on the list, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");
            }
            #endregion UpdateContentType
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with incorrect listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC61_UpdateContentType_IncorrectListName()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2766, this.Site), @"Test is executed only when R2766Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region Invoke GetListContentTypes to get the old version

            GetListContentTypesResponseGetListContentTypesResult result = this.listswsAdapter.GetListContentTypes(listName, null);

            this.Site.Assert.IsNotNull(result, "GetListContentTypes should succeed.");
            #endregion

            #region Abstract Value Bind

            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(Common.GetConfigurationPropertyValue("ListFieldText", this.Site));

            #endregion

            #region UpdateContentType
            string invalidList = DateTime.Now.TimeOfDay.ToString();
            bool caughtSoapException = false;
            try
            {
                this.listswsAdapter.UpdateContentType(invalidList, result.ContentTypes.ContentType[0].ID, properties, addFields, null, null, bool.TrueString);
                Site.Assert.Fail("UpdateContentType should fail when the list name is neither the exist GUID nor title, but it passed.");
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    2766,
                    @"[In UpdateContentType operation] Implementation does return a SOAP fault with error code 0x82000006,if listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");

                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    exp.Detail.LastChild.InnerText,
                    1975,
                    @"[In UpdateContentType operation] [If the specified listName does not correspond to a list from either of these checks, a SOAP fault should<73> be returned with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            }
            finally
            {
                Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'UpdateContentType' with incorrect listName.");
            }

            #endregion UpdateContentType
        }

        /// <summary>
        /// This test case is used to verify UpdateContentType operation in Windows SharePoint Services 3.0
        /// when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC62_UpdateContentType_InvalidListName_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2496, this.Site), @"Test is executed only when R2496Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Identify whether the soap fault is returned.
            bool isReturnSopeFault = false;
            try
            {
                // Call UpdateContentType method to update Content Type operation.
                this.listswsAdapter.UpdateContentType(
                                                        invalidListName,
                                                        string.Empty,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        string.Empty);
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2496
                bool isVerifyR2496 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2496 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText.Equals("Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).");

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2496 should be covered.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2496,
                    2496,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<73> Section 3.1.4.27: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");

                isReturnSopeFault = true;
            }

            Site.Assert.IsTrue(isReturnSopeFault, "The soap fault should be returned by the server.");
        }

        /// <summary>
        /// This test case is used to verify the negative status of updating a new field but success of adding, updating and deleting of other existing fields in UpdateContentType operation.
        /// </summary>>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC63_UpdateContentType_UpdateFieldsError()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC62_UpdateContentType_UpdateFieldsError, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldNameB = TestSuiteHelper.GetUniqueFieldName();
            string fieldNameC = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            // Call GetListContentType operation to get the content type.
            this.listswsAdapter.GetListContentType(
                                           listName,
                                           contentTypeId);

            // Construct one valid added field
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(fieldNameC);

            // Construct two update fields, one valid and the other one is invalid
            AddOrUpdateFieldsDefinition updateFields = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameA, fieldNameB);

            // Construct one valid delete field
            DeleteFieldsDefinition deleteFields = TestSuiteHelper.CreateDeleteContentTypeFields(fieldNameA);

            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();

            // Call UpdateContentType operation to update the content type.
            UpdateContentTypeResponseUpdateContentTypeResult updateResult = this.listswsAdapter.UpdateContentType(
                                                    listName,
                                                    contentTypeId,
                                                    updateProperties,
                                                    addFields,
                                                    updateFields,
                                                    deleteFields,
                                                    Convert.ToString(true));

            Site.Assert.IsNotNull(updateResult, "The result of UpdateContentType operation should not be NULL");

            bool isCaptured = updateResult.Results.Method[2].ErrorCode == "0x82000007";

            // Capture R811 if the error code is "0x82000007".
            Site.CaptureRequirementIfIsTrue(
                isCaptured,
                811,
                @"[In UpdateContentType operation] [The protocol server updates the appropriate fields on the content type.] If the fields to be updated do not exist in the specified contentType or they do not exist in the specified list, the protocol server MUST return a Response with error code 0x82000007. ");

            // Call GetListContentType operation to get the content type.
            GetListContentTypeResponseGetListContentTypeResult getResult = this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);

            Site.Assert.IsNotNull(getResult, "The result of GetListContentType operation should not be NULL");
        }

        /// <summary>
        /// The test case is used to verify that the protocol server returns a SOAP fault when there is an error on updating a property in UpdateContentType operation. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC64_UpdateContentType_UpdatePropertiesError()
        {
            #region Create a new list

            bool isCreateListSuccess = false;
            string listName = TestSuiteHelper.GetUniqueListName();
            string listGuid = TestSuiteHelper.CreateList(listName);
            isCreateListSuccess = !string.IsNullOrEmpty(listGuid);
            Site.Assert.IsTrue(isCreateListSuccess, "MSLISTSWS_S02_TC63_UpdateContentType_UpdatePropertiesError, create a new list successfully.");
            #endregion

            #region Create an content type on the list
            string contentTypeId = null;
            string fieldNameA = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldNameC = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);
            string contentTypeDisplayName = TestSuiteHelper.GetUniqueContentTypeName();

            contentTypeId = TestSuiteHelper.CreateContentType(listName, contentTypeDisplayName, new List<string> { fieldNameA });
            Site.Assert.IsNotNull(contentTypeId, "Call CreateContentType operation to create a content type successfully.");
            #endregion

            // Call GetListContentType operation to get the content type.
            this.listswsAdapter.GetListContentType(
                                            listName,
                                            contentTypeId);

            // Construct one valid added field
            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(fieldNameC);

            // Construct one valid updated field
            AddOrUpdateFieldsDefinition updateFields = TestSuiteHelper.CreateUpdateContentTypeFields(fieldNameA);

            // Construct one valid delete field
            DeleteFieldsDefinition deleteFields = TestSuiteHelper.CreateDeleteContentTypeFields(fieldNameA);

            // Construct invalid title property which contains white space
            UpdateContentTypeContentTypeProperties updateProperties = new UpdateContentTypeContentTypeProperties();
            updateProperties.ContentType = new ContentTypePropertyDefinition();

            bool isUpdateContentTypeFail = false;
            try
            {
                // Set an empty for the content type's title so that it would generate a SOAP fault
                updateProperties.ContentType.Title = string.Empty;

                // Call UpdateContentType operation to update the content type.
                this.listswsAdapter.UpdateContentType(
                                                         listName,
                                                         contentTypeId,
                                                         updateProperties,
                                                         addFields,
                                                         updateFields,
                                                         deleteFields,
                                                         Convert.ToString(true));
            }
            catch (SoapException)
            {
                isUpdateContentTypeFail = true;
            }

            // If the server return a soap fault, capture R816.
            Site.CaptureRequirementIfIsTrue(
                isUpdateContentTypeFail,
                816,
                @"[In UpdateContentType operation] [The protocol server updates the content type with the given properties.] If an error is encountered during this operation, the protocol server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to test UpdateContentType with correct listName with GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC65_UpdateContentType_WithGuid_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list should succeed.");

            #endregion Add a list

            #region Invoke GetListContentTypes to get the old version

            GetListContentTypesResponseGetListContentTypesResult result = this.listswsAdapter.GetListContentTypes(listName, null);

            this.Site.Assert.IsNotNull(result, "Calling GetListContentTypes successfully.");
            #endregion

            #region Abstract Value Bind
            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(
                                                                        Common.GetConfigurationPropertyValue("ListFieldText", this.Site));

            addFields.Fields[0].Field.Node = "//" + listId + "/" + result.ContentTypes.ContentType[0].ID + "/" + addFields.Fields[0].ID;
            #endregion

            #region UpdateContentType
            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResultUpdate
                = this.listswsAdapter.UpdateContentType(listId, result.ContentTypes.ContentType[0].ID, properties, addFields, null, null, bool.TrueString);

            if (updateContentTypeResultUpdate == null)
            {
                Site.Assert.Fail("Update contentType operation failed as Updatecontent type result returned null.");
            }

            if (Common.IsRequirementEnabled(7921, this.Site))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R7921.
                // If updating content type is successful, it means implementation does support this UpdateContentType method. R7921 can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    updateContentTypeResultUpdate,
                    7921,
                    @"Implementation does support this method[UpdateContentType]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // Verify R797.
            Site.CaptureRequirement(
                797,
                @"[In UpdateContentType operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            // Verify R828.
            Site.CaptureRequirement(
                828,
                @"[In UpdateContentType operation] [In UpdateContentTypeResponse element] [In UpdateContentTypeResult element] If no error conditions, as specified previously, cause the protocol server to return a SOAP fault, an UpdateContentTypeResult MUST be returned.");

            #endregion

            #region GetListContentType

            GetListContentTypeResponseGetListContentTypeResult getListContentType = this.listswsAdapter.GetListContentType(listId, result.ContentTypes.ContentType[0].ID);

            // Verify whether the ContentTypeFields had been updated.
            bool appropriateIsUpdate = false;

            for (int i = 0; i < getListContentType.ContentType.Fields.Items.Length; i++)
            {
                FieldDefinitionCTRemove fieldRefDefCT = (FieldDefinitionCTRemove)getListContentType.ContentType.Fields.Items[i];

                if (fieldRefDefCT.DisplayName.Equals(Common.GetConfigurationPropertyValue("ListFieldText", this.Site)))
                {
                    appropriateIsUpdate = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                appropriateIsUpdate,
                2274,
                @"[In UpdateContentType operation] The protocol server updates the appropriate fields on the content type.");

            Site.CaptureRequirementIfIsTrue(
                appropriateIsUpdate,
                2275,
                @"[In UpdateContentType operation] The protocol server updates the content type with the given properties.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateContentType operation with correct listName with title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC66_UpdateContentType_WithListTitle_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Create list should succeed.");

            #endregion

            #region Create a content type

            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string parentTypeID = TestSuiteHelper.GetContentTypeId(listId, TestSuiteHelper.GetFirstExistContentTypeName(listId));

            // Create a contentType property
            CreateContentTypeContentTypeProperties addProperties = new CreateContentTypeContentTypeProperties();
            addProperties.ContentType = new ContentTypePropertyDefinition();
            addProperties.ContentType.Description = displayName;
            addProperties.ContentType.Title = displayName;
            addProperties.ContentType.Group = displayName;

            string newContentTypeId = this.listswsAdapter.CreateContentType(
                                                    listName,
                                                    displayName,
                                                    parentTypeID,
                                                    null,
                                                    addProperties,
                                                    "TRUE");

            this.Site.Assert.IsNotNull(newContentTypeId, "Create list content type should succeed.");

            #endregion

            #region Invoke GetListContentTypes to get the old version

            GetListContentTypeResponseGetListContentTypeResult getListContentTypeResult = null;
            getListContentTypeResult = this.listswsAdapter.GetListContentType(listName, newContentTypeId);

            this.Site.Assert.IsNotNull(getListContentTypeResult, "GetListContentType operation should succeed.");

            #endregion

            #region Abstract Value Bind

            UpdateContentTypeContentTypeProperties properties = new UpdateContentTypeContentTypeProperties();
            properties.ContentType = new ContentTypePropertyDefinition();

            AddOrUpdateFieldsDefinition addFields = TestSuiteHelper.CreateAddContentTypeFields(Common.GetConfigurationPropertyValue("ListFieldText", this.Site));

            #endregion

            #region UpdateContentType

            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResultUpdate
                = this.listswsAdapter.UpdateContentType(listName, newContentTypeId, properties, addFields, null, null, bool.TrueString);

            this.Site.Assert.IsNotNull(updateContentTypeResultUpdate, "UpdateContentType operation should succeed.");

            #endregion UpdateContentType

            // Verify R798.
            Site.CaptureRequirement(
                798,
                @"[In UpdateContentType operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and, if so, use that list.");

            // Verify R799
            Site.CaptureRequirement(
                799,
                @"[In UpdateContentType operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and, if so, use that list.");

            // Verify R1999
            Site.CaptureRequirementIfAreEqual<string>(
                getListContentTypeResult.ContentType.Description,
                updateContentTypeResultUpdate.Results.ListProperties.Description,
                1999,
                @"[UpdateContentTypeResponse]ListProperties.Description: The description of the content type.");

            // Verify R2001
            Site.CaptureRequirementIfAreEqual<string>(
                getListContentTypeResult.ContentType.Group,
                updateContentTypeResultUpdate.Results.ListProperties.Group,
                2001,
                @"[UpdateContentTypeResponse]ListProperties.Group: The name of the content type group in which this content type is contained.");

            // Verify R2003
            Site.CaptureRequirementIfAreEqual<string>(
                getListContentTypeResult.ContentType.ID,
                updateContentTypeResultUpdate.Results.ListProperties.ID,
                2003,
                @"[UpdateContentTypeResponse]ListProperties.ID: The content type identifier for this content type.");

            // Verify R2004
            Site.CaptureRequirementIfAreEqual<string>(
                getListContentTypeResult.ContentType.Name,
                updateContentTypeResultUpdate.Results.ListProperties.Name,
                2004,
                @"[UpdateContentTypeResponse]ListProperties.Name: The name of the content type.");
        }

        #endregion

        #region UpdateContentTypesXmlDocument

        /// <summary>
        /// This test case is used to test UpdateContentTypesXmlDocument with incorrect contentTypeID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC67_UpdateContentTypesXmlDocument_IncorrectContentTypeId()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);
            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextTypeId = TestSuiteHelper.CreateContentType(listId, displayName, fields);
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextTypeId), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypesXmlDocument
            // Construct the newDocument parameter for the invocation below.
            UpdateContentTypesXmlDocumentNewDocument docForInvalidContentTypeId = new UpdateContentTypesXmlDocumentNewDocument();

            docForInvalidContentTypeId.ContentTypes = new UpdateContentTypesXmlDocumentNewDocumentContentType[1];

            string namespaceUri1 = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXml = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);
            string qualifiedName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            docForInvalidContentTypeId.ContentTypes[0] = new UpdateContentTypesXmlDocumentNewDocumentContentType();
            XmlDocument doc = new XmlDocument();
            docForInvalidContentTypeId.ContentTypes[0].Any = doc.CreateElement(qualifiedName, namespaceUri1);
            docForInvalidContentTypeId.ContentTypes[0].Any.InnerXml = innerXml;

            // Set an invalid ContentTypeId.
            string invalidContentTypeId = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            docForInvalidContentTypeId.ContentTypes[0].ID = invalidContentTypeId;

            try
            {
                this.listswsAdapter.UpdateContentTypesXmlDocument(listId, docForInvalidContentTypeId);
                Site.Assert.Fail("Calling UpdateContentType should fail");
            }
            catch (SoapException ex)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(ex);

                Site.CaptureRequirementIfIsNotNull(
                    ex,
                    841,
                    @"[In UpdateContentTypesXmlDocument operation] If the content type specified by the ContentType.ID attribute is not found, the protocol server MUST return a SOAP fault.");

                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    2009,
                    @"[In UpdateContentTypesXmlDocument operation] [If the content type specified by the ContentType.ID attribute is not found, the protocol server MUST return a SOAP fault.] There is no error code for this fault.");
            }

            // Construct another newDocument parameter to verify R842 and R2010.
            UpdateContentTypesXmlDocumentNewDocument docForEmptyChildElements = new UpdateContentTypesXmlDocumentNewDocument();
            docForEmptyChildElements.ContentTypes = new UpdateContentTypesXmlDocumentNewDocumentContentType[1];
            docForEmptyChildElements.ContentTypes[0] = new UpdateContentTypesXmlDocumentNewDocumentContentType();
            docForEmptyChildElements.ContentTypes[0].ID = contextTypeId;

            try
            {
                this.listswsAdapter.UpdateContentTypesXmlDocument(listId, docForEmptyChildElements);
                Site.Assert.Fail("Calling UpdateContentType should fail");
            }
            catch (SoapException ex)
            {
                string errorCode = TestSuiteHelper.GetErrorCode(ex);

                Site.CaptureRequirementIfIsNotNull(
                    ex,
                    842,
                    @"[In UpdateContentTypesXmlDocument operation] If the content type specified by the ContentType.ID attribute does not contain any child elements, the protocol server MUST return a SOAP fault.");

                Site.CaptureRequirementIfIsNull(
                    errorCode,
                    2010,
                    @"[In UpdateContentTypesXmlDocument operation] [If the content type specified by the ContentType.ID attribute does not contain any child elements, the protocol server MUST return a SOAP fault.] There is no error code for this fault.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to test UpdateContentTypesXmlDocument with incorrect listName.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC68_UpdateContentTypesXmlDocument_IncorrectListName()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(2767, this.Site) || Common.IsRequirementEnabled(2499, this.Site),
                @"Test is executed only when R2767Enabled is set to true or R2499Enabled is set to true.");

            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypesXmlDocument
            // Construct the newDocument parameter for the invocation below.
            UpdateContentTypesXmlDocumentNewDocument newDocument = new UpdateContentTypesXmlDocumentNewDocument();
            newDocument.ContentTypes = new UpdateContentTypesXmlDocumentNewDocumentContentType[1];
            string qualifiedNameOne = TestSuiteHelper.GetUniqueFieldName();
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);

            newDocument.ContentTypes[0] = new UpdateContentTypesXmlDocumentNewDocumentContentType();
            newDocument.ContentTypes[0].ID = contextType;

            XmlDocument doc = new XmlDocument();

            newDocument.ContentTypes[0].Any = doc.CreateElement(qualifiedNameOne, namespaceURIOne);
            newDocument.ContentTypes[0].Any.InnerXml = innerXmlOne;

            string invalidList = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool issoapFaultGenerated = false;
            string errorCode = string.Empty;
            string errorString = string.Empty;
            try
            {
                this.listswsAdapter.UpdateContentTypesXmlDocument(invalidList, newDocument);

                Site.Assert.Fail("Calling UpdateContentType should fail");
            }
            catch (SoapException exp)
            {
                issoapFaultGenerated = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                errorString = TestSuiteHelper.GetErrorString(exp);
            }

            this.Site.Assert.IsTrue(issoapFaultGenerated, "There should be a soap fault generated when call UpdateContentTypesXmlDocument operation with invalidListName");

            if (Common.IsRequirementEnabled(2767, this.Site))
            {
                this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value:errorCode[{0}] for requirement #R2767 #R2011",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    "0x82000006".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                    2767,
                    @"[In UpdateContentTypesXmlDocument operation] Implementation does return a SOAP fault with error code 0x82000006, if the list specified by the listName is not found. (SharePoint Foundation 2010 and above follow this behavior.)");

                Site.CaptureRequirementIfIsTrue(
                        "0x82000006".Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                      2011,
                      @"[In UpdateContentTypesXmlDocument operation] [If the list specified by the listName is not found, the protocol server should<74> return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user.");
            }

            if (Common.IsRequirementEnabled(2499, this.Site))
            {
                this.Site.Log.Add(
                       LogEntryKind.Debug,
                       "The actual value:errorCode[{0}] errorString[{1}] for requirement #R2499",
                       string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode,
                        string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

                string expectedErrorString = "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
                Site.CaptureRequirementIfIsTrue(
                      string.IsNullOrEmpty(errorCode) && expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase),
                      2499,
                      @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<75> Section 3.1.4.28: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateContentTypesXmlDocument operation with correct listName GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC69_UpdateContentTypesXmlDocument_WithGuid_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypesXmlDocument
            // Construct the newDocument parameter for the invocation below.
            UpdateContentTypesXmlDocumentNewDocument newDocument = new UpdateContentTypesXmlDocumentNewDocument();
            newDocument.ContentTypes = new UpdateContentTypesXmlDocumentNewDocumentContentType[1];
            string qualifiedNameOne = TestSuiteHelper.GetUniqueFieldName();
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);

            newDocument.ContentTypes[0] = new UpdateContentTypesXmlDocumentNewDocumentContentType();
            newDocument.ContentTypes[0].ID = contextType;

            XmlDocument doc = new XmlDocument();

            newDocument.ContentTypes[0].Any = doc.CreateElement(qualifiedNameOne, namespaceURIOne);
            newDocument.ContentTypes[0].Any.InnerXml = innerXmlOne;

            UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult updateContentTypesXmlDocumentResult
                = this.listswsAdapter.UpdateContentTypesXmlDocument(listId, newDocument);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R8351.
            // If updating content types XML document is successful, it means implementation does support this UpdateContentTypesXmlDocument method. R8351 can be captured.
            if (Common.IsRequirementEnabled(8351, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    updateContentTypesXmlDocumentResult,
                    8351,
                    @"Implementation does support this method[UpdateContentTypesXmlDocument]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If protocol SUT returns the response without any Soap Fault, then capture R843
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                843,
                @"[In UpdateContentTypesXmlDocument operation] If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateContentTypesXmlDocument operation with listName title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S02_TC70_UpdateContentTypesXmlDocument_WithListTitle_Succeed()
        {
            #region Add a list
            string listName = TestSuiteHelper.GetUniqueListName();

            string listId = TestSuiteHelper.CreateList(listName);

            bool isListCreatedSuccessful = !string.IsNullOrEmpty(listId);
            this.Site.Assert.IsTrue(isListCreatedSuccessful, "Creating list successfully.");

            #endregion Add a list

            #region CreateContentType
            string displayName = TestSuiteHelper.GetUniqueContentTypeName();
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<string> fields = new List<string> { fieldName };

            string contextType = TestSuiteHelper.CreateContentType(listId, displayName, fields);
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(contextType), "Calling CreateContentType successfully.");
            #endregion CreateContentType

            #region  UpdateContentTypesXmlDocument
            // Construct the newDocument parameter for the invocation below.
            UpdateContentTypesXmlDocumentNewDocument newDocument = new UpdateContentTypesXmlDocumentNewDocument();
            newDocument.ContentTypes = new UpdateContentTypesXmlDocumentNewDocumentContentType[1];
            string qualifiedNameOne = TestSuiteHelper.GetUniqueFieldName();
            string namespaceURIOne = Common.GetConfigurationPropertyValue("ContentTypeXmlNamespaceUri", this.Site);
            string innerXmlOne = Common.GetConfigurationPropertyValue("ValidContentTypeXmlDocument", this.Site);

            newDocument.ContentTypes[0] = new UpdateContentTypesXmlDocumentNewDocumentContentType();
            newDocument.ContentTypes[0].ID = contextType;

            XmlDocument doc = new XmlDocument();

            newDocument.ContentTypes[0].Any = doc.CreateElement(qualifiedNameOne, namespaceURIOne);
            newDocument.ContentTypes[0].Any.InnerXml = innerXmlOne;

            UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult updateContentTypesXmlDocumentResult
                = this.listswsAdapter.UpdateContentTypesXmlDocument(listName, newDocument);

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                857,
                @"[In UpdateContentTypesXmlDocument operation] [In UpdateContentTypesXmlDocumentResponse element] [In UpdateContentTypesXmlDocumentResult element] If no error conditions cause the protocol server to return a SOAP fault, an UpdateContentTypesXmlDocumentResult MUST be returned.");

            // If the list whose list title corresponds to the specified listName exists and it is updated successfully, the server will return a successful response.
            // So if the response of the operation isn't null, then capture R844, R845, R848 and R2278.
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                844,
                @"[In UpdateContentTypesXmlDocument operation] If the specified listName is not a valid GUID, check if the listName corresponds to the list title of a list on the site and, if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                845,
                @"[In UpdateContentTypesXmlDocument operation] If the specified listName does not correspond to the identification of a list on the site, check if the listName corresponds to the list title of a list on the site and, if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                848,
                @"[In UpdateContentTypesXmlDocument operation] Otherwise [If the newDocument follows the schema, the content type specified by the ContentType.ID attribute contains child elements, the specified listName corresponds to the identification or the list title of a list on the site and there is no XML document with the same namespace URI exist], the new XML Document MUST be added to the XML document collection that the content type specified by the ContentType.ID attribute.");

            Site.CaptureRequirementIfIsNotNull(
                updateContentTypesXmlDocumentResult,
                2278,
                @"[In UpdateContentTypesXmlDocument operation] Otherwise [If the newDocument follows the schema, the content type specified by the ContentType.ID attribute contains child elements, the specified listName corresponds to the identification or the list title of a list on the site and there is no XML document with the same namespace URI exist] the protocol server MUST return a success UpdateContentTypesXmlDocumentResult.");

            #endregion
        }

        #endregion

        #endregion

        #region Override methods

        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.listswsAdapter = this.Site.GetAdapter<IMS_LISTSWSAdapter>();

            Common.CheckCommonProperties(this.Site, true);

            #region initialization
            if (!TestSuiteHelper.GuardEnviromentClean())
            {
                Site.Debug.Fail("The test environment is not clean, refer the log files for details.");
            }

            // Initialize the TestSuiteHelper
            TestSuiteHelper.Initialize(this.Site, this.listswsAdapter);
            #endregion
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            #region new clean up
            TestSuiteHelper.CleanUp();
            #endregion
        }

        #endregion
    }
}