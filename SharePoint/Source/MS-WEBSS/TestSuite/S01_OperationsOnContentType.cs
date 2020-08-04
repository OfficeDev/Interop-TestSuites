namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The Test Suite of MS-WEBSS. Protocol client tries to perform operations associated with content types.
    /// </summary>
    [TestClass]
    public class S01_OperationsOnContentType : TestSuiteBase
    {
        #region Additional test attributes, initialization and clean up

        /// <summary>
        /// Class initialization.
        /// </summary>     
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        /// <summary>
        /// This test case aims to verify create a new content type on the context site with valid parameters.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC01_CreateContentType()
        {
            // Create a new content type on the context site.
            ContentTypeType.ContentType = new ContentTypePropertyDefinition();
            ContentTypeType.ContentType.Description = this.ContentTypeDescription;
            ContentTypeType.ContentType.Title = this.ContentTypeTypeTitle;
            AddOrUpdateFieldsDefinition fields = GenerateNewFields();
            string contentTypeId = Adapter.CreateContentType(DisplayName, Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site), fields, ContentTypeType);

            // Verify MS-WEBSS requirement: MS-WEBSS_R712, MS-WEBSS_R1023
            // Captured by Schema.
            Site.CaptureRequirementIfIsNotNull(
                contentTypeId,
                712,
                @"[In CreateContentTypeResponse] CreateContentTypeResult: If no error conditions as specified in the preceding section cause the protocol server to return a SOAP exception, a CreateContentTypeResult MUST be returned containing the content type identifier of the content type created.");

            if (Common.IsRequirementEnabled(1023, this.Site))
            {
                if (contentTypeId != null)
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1023
                    // When the system under test product name is Windows SharePoint Services 3.0 and above, if the server returns a
                    //  positive information when invoke CreateContentType operation, then the requirement can be captured.
                    Site.CaptureRequirement(
                        1023,
                        @"[In Appendix B: Product Behavior] Implementation does support this[CreateContentType] operation. (<9> Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }

            // Create a new content type, set newFields to null.
            CreateContentTypeContentTypeProperties contentTypeProperties = new CreateContentTypeContentTypeProperties();
            contentTypeProperties.ContentType = new ContentTypePropertyDefinition();
            contentTypeProperties.ContentType.Description = this.GenerateRandomString(10);
            contentTypeProperties.ContentType.Title = this.GenerateRandomString(10);
            Adapter.CreateContentType(this.GenerateRandomString(10), Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site), null, contentTypeProperties); 
        }

        /// <summary>
        /// This test case aims to verify remove a given content type from the site with valid parameters.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC02_DeleteContentType()
        {
            // Create a new content type on the context site.
            ContentTypeType.ContentType = new ContentTypePropertyDefinition();
            ContentTypeType.ContentType.Description = this.ContentTypeDescription;
            ContentTypeType.ContentType.Title = this.ContentTypeTypeTitle;
            AddOrUpdateFieldsDefinition fields = GenerateNewFields();
            string contentTypeID = Adapter.CreateContentType(DisplayName, Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site), fields, ContentTypeType);

            // Remove a specified content type from the site.
            DeleteContentTypeResponseDeleteContentTypeResult deleteResult = Adapter.DeleteContentType(contentTypeID);

            Site.Assert.IsNotNull(deleteResult.Success, "The value of Success in DeleteContentTypeResult should not be null");
            if (Common.IsRequirementEnabled(1025, this.Site))
            {
                if (deleteResult.Success != null)
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1025
                    // When the system under test product name is Windows SharePoint Services 3.0 and above, if the server returns a
                    //  positive information when invoke DeleteContentType operation, then the requirement can be captured.
                    this.Site.CaptureRequirement(
                        1025,
                        @"[In Appendix B: Product Behavior] Implementation does support this[DeleteContentType] operation. (<12>Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }

            // Retrieve all content types for a specified context site.
            GetContentTypesResponseGetContentTypesResult getContentTypesResult = Adapter.GetContentTypes();
            bool isDelContentType = false;
            Site.Assert.IsNotNull(getContentTypesResult.ContentTypes, "The value of ContentTypes in GetContentTypesResponse should not be null");
            for (int i = 0; i < getContentTypesResult.ContentTypes.Length; i++)
            {
                if (getContentTypesResult.ContentTypes[i].ID != contentTypeID)
                {
                    isDelContentType = true;
                    break;
                }
            }

            // If this contentTypeID not exist in the response of GetContentTypes,
            // contentTypeID has been removed by operation DeleteContentType.
            // Verify MS-WEBSS requirement: MS-WEBSS_R109
            Site.CaptureRequirementIfIsTrue(
                isDelContentType,
                109,
                @"[In DeleteContentType] If no exceptions are encountered during the processing of the request, the content type with the requested identifier MUST be removed from the list of content types on the site (2).");
        }

        /// <summary>
        /// This test case aims to verify the GetContentTypes operation, and the response should contain all content types on a site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC03_GetContentTypesWithElements()
        {
            // Obtain content types data for a specified content types identifier.
            GetContentTypesResponseGetContentTypesResult contentTypeResult = Adapter.GetContentTypes();

            // Verify MS-WEBSS requirement: MS-WEBSS_R241
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                contentTypeResult != null,
                241,
                @"[In GetContentTypesResponse] GetContentTypesResult: Container element for the rest of the results.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R242
            // If the content types is not null, the R242 will be captured.
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                contentTypeResult.ContentTypes != null,
                242,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes: Container for each block of content type data.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R243
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                contentTypeResult.ContentTypes[0] != null,
                243,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType: Instance of content type data.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R244
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                contentTypeResult.ContentTypes[0].Name != null,
                244,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.Name: Name of the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R245
            bool isVerifyR245 = false;
            if (contentTypeResult.ContentTypes[0].ID != null)
            {
                isVerifyR245 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR245,
                245,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.ID: Content type identifier for this[GetContentTypesResult.ContentTypes.ContentType.ID] content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R246
            bool isVerifyR246 = false;
            if (contentTypeResult.ContentTypes[0].Description != null)
            {
                isVerifyR246 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR246,
                246,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.Description: Description of the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R247
            bool isVerifyR247 = false;
            if (contentTypeResult.ContentTypes[0].Group != null)
            {
                isVerifyR247 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR247,
                247,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.Group: Name of the content type group that contains this[GetContentTypesResult.ContentTypes.ContentType.Group] content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R250 
            bool isVerifyR250 = false;
            int len = contentTypeResult.ContentTypes.Length;
            for (int i = 0; i < len; i++)
            {
                string siteCollectionUrl = Common.GetConfigurationPropertyValue("SiteCollectionUrl", this.Site);

                siteCollectionUrl = siteCollectionUrl.ToLower();
                if (contentTypeResult.ContentTypes[i].Scope.ToLower().Contains(siteCollectionUrl))
                {
                    isVerifyR250 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR250,
                250,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.Scope: Fully qualified URL of the content type scope.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R251
            bool isVerifyR251 = false;
            for (int i = 0; i < len; i++)
            {
                // The type of version is int, and the version number is greater than or equal to zero. If the condition is established,
                // the requirement will be captured.
                int versionNum = contentTypeResult.ContentTypes[i].Version;
                Site.Assert.IsFalse(versionNum < 0, "Failed to be returned the version for the GetContentTypes operation.");
                if (versionNum >= 0)
                {
                    isVerifyR251 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR251,
                251,
                @"[In GetContentTypesResponse] GetContentTypesResult.ContentTypes.ContentType.Version: Specifies the current version of the content type.");
        }

        /// <summary>
        /// This test case aims to verify GetContentTypes operation with valid content in the response.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC04_GetContentTypes()
        {
            #region Prepare for UpdateContentType.

            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);
            this.UpdateFields = this.GenerateUpdateFields();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            NewFields.Fields[0].ID = this.NewFieldsID;
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();

            #endregion

            #region Get version before updating content type.

            int versionUpdateBefore = -1;
            int versionUpdateAfter = -1;

            // Retrieve all content types for a specified context site.
            GetContentTypesResponseGetContentTypesResult result = Adapter.GetContentTypes();

            for (int i = 0; i < result.ContentTypes.Length; i++)
            {
                if (result.ContentTypes[i].Name == this.ContentTypeTypeTitle)
                {
                    versionUpdateBefore = result.ContentTypes[i].Version;
                    break;
                }
            }

            #endregion

            #region Get version after updating content type.

            // Update a content type on the context site.
            Adapter.UpdateContentType(contentTypeID, newProperty, this.NewFields, this.UpdateFields, this.DeleteField);

            // Retrieve all content types for a specified context site.
            result = Adapter.GetContentTypes();
            for (int i = 0; i < result.ContentTypes.Length; i++)
            {
                if (result.ContentTypes[i].Name == this.ContentTypeTypeTitle)
                {
                    versionUpdateAfter = result.ContentTypes[i].Version;
                    break;
                }
            }

            #endregion

            Site.CaptureRequirementIfAreEqual<bool>(
                 true,
                 versionUpdateAfter > versionUpdateBefore,
                 252,
                 @"[In GetContentTypesResponse] The protocol server MUST increment the value each time the content type definition is edited.");
        }

        /// <summary>
        /// This test case aims to verify UpdateContentType operation with invalid delete field.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC05_UpdateContentTypeWithInvalidDeleteField()
        {
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);

            this.UpdateFields = this.GenerateUpdateFields();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();
            DeleteField.Fields[0].ID = Guid.NewGuid().ToString();
            DeleteField.Fields[0].Field.Name = this.GenerateRandomString(10);

            // Update a content type on the context site.
            UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResult = Adapter.UpdateContentType(contentTypeID, newProperty, NewFields, UpdateFields, DeleteField);

            // Verify MS-WEBSS requirement: MS-WEBSS_R561
            Site.CaptureRequirementIfAreNotEqual<string>(
                SoapErrorCode.ErrorCode0x00000000,
                updateContentTypeResult.Results.Method[0].ErrorCode,
                561,
               @"[In UpdateContentTypeResponse] Otherwise[if a reference to the field is not found on the content type], the value of ErrorCode MUST be nonzero.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R564
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeResult.Results.Method[0].ErrorText,
                564,
                @"[In UpdateContentTypeResponse] ErrorText: If ErrorCode is nonzero, ErrorText MUST contain an error string.");
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentType operation when the updateFields can be found.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC06_UpdateContentTypeFoundTheField()
        {
            #region Prepare for UpdateContentType.
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);
            this.UpdateFields = this.GenerateUpdateFields();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            NewFields.Fields[0].ID = this.NewFieldsID;
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();

            #endregion

            // Update a content type on the context site.
            UpdateContentTypeResponseUpdateContentTypeResult updateContentType = Adapter.UpdateContentType(contentTypeID, newProperty, NewFields, UpdateFields, DeleteField);

            #region Capture requirements.

            // Verify MS-WEBSS requirement: MS-WEBSS_R570
            Site.CaptureRequirementIfAreEqual<string>(
                contentTypeID,
                updateContentType.Results.ListProperties.ID,
                570,
                @"[In UpdateContentTypeResponse] ListProperties.ID: Content type identifier for this[ListProperties.ID] content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R571
            Site.CaptureRequirementIfAreEqual<ONEONLY>(
                ONEONLY.Item1,
                updateContentType.Results.ListProperties.Locs,
                571,
                @"[In UpdateContentTypeResponse] ListProperties.Locs: Specifies whether the content type has localizations.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R573
            Site.CaptureRequirementIfAreEqual<string>(
                this.ContentTypeTypeTitle,
                updateContentType.Results.ListProperties.Name,
                573,
                @"[In UpdateContentTypeResponse] ListProperties.Name: Localizations for the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R576
            Site.CaptureRequirementIfAreEqual<TRUEONLY>(
                TRUEONLY.TRUE,
                updateContentType.Results.ListProperties.ReadOnly,
                576,
                @"[In UpdateContentTypeResponse] ListProperties.ReadOnly: Specifies whether the content type can be edited.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R577
            Site.CaptureRequirementIfAreEqual<FALSEONLY>(
                FALSEONLY.FALSE,
                updateContentType.Results.ListProperties.RequireClientRenderingOnNew,
                577,
                @"[In UpdateContentTypeResponse] ListProperties.RequireClientRenderingOnNew: Specifies whether the browser uses the ActiveX object to create a new item.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R578
            Site.CaptureRequirementIfAreEqual<TRUEONLY>(
                TRUEONLY.TRUE,
                updateContentType.Results.ListProperties.Sealed,
                578,
                @"[In UpdateContentTypeResponse] ListProperties.Sealed: Specifies whether another content type can inherit from the current content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R580
            bool isVerifyR580 = false;
            if (updateContentType.Results.ListProperties.Version >= 0)
            {
                isVerifyR580 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR580,
                580,
                @"[In UpdateContentTypeResponse] ListProperties.Version: Specifies the current version of the content type.");

            #endregion
        }

        /// <summary>
        /// This test case aims to verify CreateContentType operation with invalid parameters. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC07_CreateContentTypeInvalidContentTypeId()
        {
            try
            {
                // Create a new content type on the context site.
                Adapter.CreateContentType(this.DisplayName, this.GenerateRandomString(6), this.NewFields, this.ContentTypeType);
                Site.Assert.Fail("The expected SOAP fault is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R46
                Site.CaptureRequirement(
                    46,
                    @"[In CreateContentType] If the parentType  is not the content type identifier of a content type that is available on the context site, a SOAP fault MUST be returned.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R50
                Site.CaptureRequirement(
                    50,
                    @"[In CreateContentType] The protocol server MUST return a SOAP fault if a problem is encountered in any of these steps.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName longer than 128 characters.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC08_CreateContentTypeInvalidDisplayNameGreaterThan128Characters()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateRandomString(129));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R68
                Site.CaptureRequirement(
                    68,
                    @"[In CreateContentType] [A SOAP exception MUST be returned if any of the following conditions applies to the unencoded version of this element:] It[displayName] is longer than 128 characters.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName beginning with whitespace.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC09_CreateContentTypeInvalidDisplayNameBeginsWithWhiteSpace()
        {
            bool isVerifyR788 = false;
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(DisplayName.PadRight(DisplayName.Length + 1));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException exp)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R788
                if (exp != null)
                {
                    isVerifyR788 = true;
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-WEBSS_R788", "If the return the no empty soap exception, the requirement will be captured.");
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR788,
                    788,
                    @"[In CreateContentType] [A SOAP exception MUST be returned if any of the following conditions applies to the unencoded version of this element:] It[displayName] begins with whitespace.");
            }
        }

        /// <summary>
        /// This test case aims to verify the the CreateContentType operation with invalid displayName ending with a period.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC10_CreateContentTypeInvalidDisplayNameEndWithPeriod()
        {
            try
            {
                string displayName = DisplayName + (char)CONST_CHARS.Period;

                // Create a new content type on the context site.
                this.CreateContentType(displayName);
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R69
                // the displayName ends with a period or begins or ends with whitespace,
                // the length of valid name of the content type to be created, the length of name is less than 128,
                // if the SOAP exception is returned, then the following requirement will be captured.
                Site.CaptureRequirement(
                    69,
                    @"[In CreateContentType] [A SOAP exception MUST be returned if any of the following conditions applies to the unencoded version of this element:] It[displayName] ends with a period or begins or ends with whitespace.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName ending with whitespace.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC11_CreateContentTypeInvalidDisplayNameEndWithSpace()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(DisplayName.PadRight(DisplayName.Length + 1));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // The displayName ends with whitespace, the length of valid name of the content type to be created,
                // the length of name is less than 128, if the SOAP exception is returned, then the following requirement will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R789
                Site.CaptureRequirement(
                    789,
                    @"[In CreateContentType] [A SOAP exception MUST be returned if any of the following conditions applies to the unencoded version of this element:] It[displayName] ends with whitespace.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName with backslash.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC12_CreateContentTypeInvalidBackslash()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Backslash));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R790
                Site.CaptureRequirement(
                    790,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains  character '\'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName with slash.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC13_CreateContentTypeInvalidSlash()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Slash));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R791
                Site.CaptureRequirement(
                    791,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains character '/'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName with colon. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC14_CreateContentTypeInvalidColon()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Colon));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R792
                Site.CaptureRequirement(
                    792,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains  character ' :' .");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains asterisk. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC15_CreateContentTypeInvalidAsterisk()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Asterisk));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R793
                Site.CaptureRequirement(
                    793,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains  character '*'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains question mark.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC16_CreateContentTypeInvalidQuestionMark()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.QuestionMark));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R794
                Site.CaptureRequirement(
                    794,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains  character '?' .");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains quotation mark.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC17_CreateContentTypeInvalidQuotationMark()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.QuotationMark));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R795
                Site.CaptureRequirement(
                    795,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '""'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains number sign.  
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC18_CreateContentTypeInvalidNumberSign()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.NumberSign));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R796
                Site.CaptureRequirement(
                    796,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '#'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains percent sign.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC19_CreateContentTypeInvalidPercentSign()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.PercentSign));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R797
                Site.CaptureRequirement(
                    797,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '%'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains less than sign.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC20_CreateContentTypeInvalidLessThanSign()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.LessThanSign));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R798
                Site.CaptureRequirement(
                    798,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '<'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains great than sign.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC21_CreateContentTypeInvalidGreaterThanSign()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.GreaterThanSign));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R799
                Site.CaptureRequirement(
                    799,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '>'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains opening curly braces.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC22_CreateContentTypeInvalidOpeningCurlyBraces()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.OpeningCurlyBraces));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R800
                Site.CaptureRequirement(
                    800,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '{'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains ClosingCurlyBraces.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC23_CreateContentTypeInvalidClosingCurlyBraces()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.ClosingCurlyBraces));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R801
                Site.CaptureRequirement(
                    801,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '}'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains Ampersand.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC24_CreateContentTypeInvalidAmpersand()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Ampersand));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R802
                Site.CaptureRequirement(
                    802,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '&'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName with period.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC25_CreateContentTypeInvalidWithPeriod()
        {
            try
            {
                this.NewFields = this.GenerateNewFields();
                ContentTypeType.ContentType = new ContentTypePropertyDefinition();
                ContentTypeType.ContentType.Description = this.ContentTypeDescription;
                ContentTypeType.ContentType.Title = this.ContentTypeTypeTitle;

                // Create a new content type on the context site.
                Adapter.CreateContentType(this.DisplayName, this.GenerateRandomString(6), this.NewFields, this.ContentTypeType);
                Site.Assert.Fail("The expected SOAP fault is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R45
                Site.CaptureRequirement(
                    45,
                    @"[In CreateContentType] If the parentType is not a valid content type identifier, a SOAP fault MUST be returned.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName with double periods. 
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC26_CreateContentTypeInvalidDoublePeriods()
        {
            try
            {
                string displayName = DisplayName + (char)CONST_CHARS.Period + (char)CONST_CHARS.Period;

                // Create a new content type on the context site.
                this.CreateContentType(displayName);
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R805
                Site.CaptureRequirement(
                    805,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains two consecutive periods.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains tab.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC27_CreateContentTypeInvalidTab()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.Tab));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R804
                Site.CaptureRequirement(
                    804,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains a tab.");
            }
        }

        /// <summary>
        /// This test case aims to verify the CreateContentType operation with invalid displayName contains verticalBar.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC28_CreateContentTypeInvalidVerticalBar()
        {
            try
            {
                // Create a new content type on the context site.
                this.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.VerticalBar));
                Site.Assert.Fail("The expected SOAP exception is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R803
                Site.CaptureRequirement(
                    803,
                    @"[In CreateContentType] A SOAP exception MUST be returned if dispalyName  contains   character '|'.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetContentType operation with invalid contentTypeId which is a null content type ID.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC29_GetContentTypeInvalidWithNull()
        {
            try
            {
                // Obtain content type data for a specified content type identifier.
                Adapter.GetContentType(null);
                Site.Assert.Fail("The expected SOAP exception is not returned for the GetContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R184
                Site.CaptureRequirement(
                    184,
                    @"[In GetContentType] If contentTypeId is null, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetContentType operation with invalid contentTypeId which is an empty content type ID.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC30_GetContentTypeInvalidWithEmpty()
        {
            try
            {
                // Obtain content type data for a specified content type identifier.
                Adapter.GetContentType(string.Empty);
                Site.Assert.Fail("The expected SOAP exception is not returned for the GetContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R807
                Site.CaptureRequirement(
                    807,
                    @"[In GetContentType] If contentTypeId is  empty, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetContentType operation with invalid contentTypeId which does not match the content type ID of a content type on the site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC31_GetContentTypeInvalidateWithInvalidContentType()
        {
            try
            {
                // Obtain content type data for a specified content type identifier.
                Adapter.GetContentType(this.GenerateRandomString(6));
                Site.Assert.Fail("The expected SOAP exception is not returned for the GetContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R185
                Site.CaptureRequirement(
                    185,
                    @"[In GetContentType] If contentTypeId does not match the content type identifier of a content type on the site, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentType operation with invalid contentTypeId.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC32_UpdateContentTypeInvalidContentTypeId()
        {
            try
            {
                // Update a content type on the context site.
                UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
                Adapter.UpdateContentType(this.GenerateRandomString(6), newProperty, this.NewFields, this.UpdateFields, this.DeleteField);
                Site.Assert.Fail("The expected SOAP exception is not returned for the UpdateContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R526
                Site.CaptureRequirement(
                    526,
                    @"[In UpdateContentType] If the contentTypeId is not the content type identifier of a content type on the context site, the protocol server MUST return a SOAP exception.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R736
                Site.CaptureRequirement(
                    736,
                    @"[In UpdateContentType] If the contentTypeId is not a valid content type identifier, the protocol server MUST return a SOAP exception.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R531
                Site.CaptureRequirement(
                    531,
                    @"[In UpdateContentType] The protocol server MUST return a SOAP exception in the specified cases.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R531
                Site.CaptureRequirement(
                    532,
                    @"[In UpdateContentType] The protocol server finishes the operations before returning any error.");
            }
        }

        /// <summary>
        /// This test case aims to verify CreateContentType operation with invalid displayName that has existed or used.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC33_CreateContentTypeWithSameName()
        {
            // Create a new content type on the context site.
            this.CreateContentType(this.DisplayName);
            try
            {
                this.CreateContentType(this.DisplayName);
                Site.Assert.Fail("The expected SOAP fault is not returned for the CreateContentType operation.");
            }
            catch (SoapException soapException)
            {
                if (soapException != null)
                {
                    // Catch the exception, then the following requirements will be captured.
                    // Verify MS-WEBSS requirement: MS-WEBSS_R64
                    Site.CaptureRequirement(
                        64,
                        @"[In CreateContentType] A SOAP exception MUST be returned if any of the following conditions applies to the unencoded version of this element[displayName]:
 It[displayName] contains the name of an existing content type on the context site, ignoring case.");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the optional behaviors in GetContentType operation.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC34_UpdateContentTypeWithShouldMay()
        {
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            this.UpdateFields = this.GenerateUpdateFields();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();

            // Update a content type on the context site.
            UpdateContentTypeResponseUpdateContentTypeResult updateResult = Adapter.UpdateContentType(contentTypeID, newProperty, NewFields, UpdateFields, DeleteField);

            // Verify MS-WEBSS requirement: MS-WEBSS_R1038
            // When the System Under Test product name is Windows SharePoint Services 3.0 and above, if the server returns a
            //  positive information when invoke UpdateContentType operation, then the requirement can be captured.
            bool isVerifyR1038 = false;
            int len = updateResult.Results.Method.Length;
            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1038, this.Site), "This operation UpdateContentType failed.");
            if (Common.IsRequirementEnabled(1038, this.Site))
            {
                for (int i = 0; i < len; i++)
                {
                    if (updateResult.Results.Method[i].ID.Equals(NewFields.Fields[0].ID, StringComparison.OrdinalIgnoreCase)
                        || updateResult.Results.Method[i].ID.Equals(UpdateFields.Fields[0].ID, StringComparison.OrdinalIgnoreCase)
                        || updateResult.Results.Method[i].ID.Equals(DeleteField.Fields[0].ID, StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifyR1038 = true;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1038,
                    1038,
                    @"[In Appendix B: Product Behavior] Implementation does support this[UpdateContentType] operation.(<24>Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case aims to verify the DeleteContentType operation with invalid content type ID.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC35_DeleteContentTypeInvalidIdentifier()
        {
            try
            {
                // Remove a specified content type from the site.
                Adapter.DeleteContentType(string.Empty);
                Site.Assert.Fail("The expected SOAP exception is not returned for the DeleteContentTypeResponse operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R120
                Site.CaptureRequirement(
                    120,
                    @"[In DeleteContentTypeResponse] If the operation fails, a SOAP exception MUST be returned.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R106
                Site.CaptureRequirement(
                    106,
                    @"[In DeleteContentType] If the contentTypeId element is not a valid content type identifier, the protocol server MUST return a SOAP exception.");
            }

            try
            {
                // Remove a specified content type from the site.
                Adapter.DeleteContentType(this.GenerateRandomString(6));
                Site.Assert.Fail("The expected  SOAP exception is not returned for the DeleteContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R107
                Site.CaptureRequirement(
                    107,
                    @"[In DeleteContentType] If the content type is not found, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the DeleteContentType operation with valid contentTypeId which is a Read-Only content type.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC36_DeleteContentTypeReadOnly()
        {
            try
            {
                // Create a new content type on the context site.
                string contentTypeID = CreateContentType(DisplayName);

                // Set The content type is read-only
                SutAdapter.SetContentTypeReadOnlyOrSealed(Common.GetConfigurationPropertyValue("webSiteName", this.Site), this.ContentTypeTypeTitle, true, false);

                // Remove a specified content type from the site.
                Adapter.DeleteContentType(contentTypeID);
                Site.Assert.Fail("The expected SOAP exception is not returned for the DeleteContentType operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R108
                Site.CaptureRequirement(
                    108,
                    @"[In DeleteContentType] If the content type is read-only or cannot be deleted for any other reason, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case is used to test CreateContentType operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC37_CreateContentType_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                this.CreateContentType(this.DisplayName);
                Site.Assert.Fail("The expected http status code is not returned for the CreateContentType operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R35
                // COMMENT: When the CreateContentType operation is invoked by unauthenticated user, 
                // if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                HttpStatusCode.Unauthorized,
                resp.StatusCode,
                35,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[CreateContentType], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case is used to test DeleteContentType operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC38_DeleteContentType_Unauthenticated()
        {
            string contentTypeID = CreateContentType(DisplayName);
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.DeleteContentType(contentTypeID);
                Site.Assert.Fail("When the DeleteContentType operation is invoked by unauthenticated user, if the server return the expected http status code, then the requirement can be captured.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1071
                // COMMENT: When the DeleteContentType operation is invoked by unauthenticated user, 
                // if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1071,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[DeleteContentType], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetContentType operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC39_GetContentType_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetContentType(Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the GetContentType operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1075
                // COMMENT: When the GetContentType operation is invoked by unauthenticated user, if 
                // the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1075,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetContentType], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetContentTypes operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC40_GetContentTypes_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetContentTypes();
                Site.Assert.Fail("The expected http status code is not returned for the GetContentTypes operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1076
                // COMMENT: When the GetContentTypes operation is invoked by unauthenticated user, if 
                // the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1076,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetContentTypes], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentTypes operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC41_UpdateContentType_Unauthenticated()
        {
            string contentTypeID = CreateContentType(DisplayName);

            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.UpdateContentType(
                    contentTypeID,
                    new UpdateContentTypeContentTypeProperties(),
                    new AddOrUpdateFieldsDefinition(),
                    new AddOrUpdateFieldsDefinition(),
                    new DeleteFieldsDefinition());
                Site.Assert.Fail("The expected http status code is not returned for the UpdateContentType operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1087
                // COMMENT: When the UpdateContentType operation is invoked by unauthenticated user, 
                // if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1087,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[UpdateContentType], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify CreateContentType operation with invalid parent type.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC42_CreateContentTypeInvalidParentType()
        {
            bool isExceptionOccuringForDispalyName = false;
            bool isExceptionOccuringForContentTypeId = false;

            // Condition 1 : Invalid contentType
            try
            {
                // Create a new content type on the context site.
                Adapter.CreateContentType(this.DisplayName, this.GenerateRandomString(6), this.NewFields, this.ContentTypeType);
                Site.Assert.Fail("The expected SOAP fault is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                isExceptionOccuringForContentTypeId = true;
            }

            // Condition 2 : Invalid Display Name
            try
            {
                // Create a new content type on the context site.
                Adapter.CreateContentType(this.GenerateInvalidDisplayName(CONST_CHARS.PercentSign), Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site), this.NewFields, this.ContentTypeType);
                Site.Assert.Fail("The expected SOAP fault is not returned for the CreateContentType operation.");
            }
            catch (SoapException)
            {
                isExceptionOccuringForDispalyName = true;
            }

            bool isR47Verified = isExceptionOccuringForContentTypeId && isExceptionOccuringForDispalyName;

            // Catch the exception, then the following requirements will be captured.
            // Verify MS-WEBSS requirement: MS-WEBSS_R47
            Site.CaptureRequirementIfIsTrue(
                isR47Verified,
                47,
                @"[In CreateContentType] [The protocol server MUST return a SOAP fault if a problem is encountered in following option.] The protocol server creates a new content type on the context site with the specified parent content type and display name.");
        }

        /// <summary>
        /// This test case aims to verify CreateContentType operation with invalid parameters.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC43_CreateContentTypeAddAndUpdateNewFields()
        {
            // Verify MS-WEBSS requirement: MS-WEBSS_R49
            // Create a new content type on the context site.
            string contentTypeInvalidID = CreateContentType(DisplayName) + "invalidId";
            this.UpdateFields = this.GenerateUpdateFields();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();
            try
            {
                // Update a content type on the context site.
                UpdateContentTypeResponseUpdateContentTypeResult actualResult = Adapter.UpdateContentType(contentTypeInvalidID, newProperty, NewFields, UpdateFields, DeleteField);
                Site.Assert.Fail("Failed to get the result of update content type {0}", actualResult);
            }
            catch (SoapException)
            {
                // If have a SOAP exception be returned from the server, this requirement will be covered.
                Site.CaptureRequirement(
                    49,
                    @"[In CreateContentType] [The protocol server MUST return a SOAP fault if a problem is encountered in following option]The protocol server updates the content type with the specified properties.");
            }
        }

        /// <summary>
        /// This test case aims to verify the operation UpdateContentType on the context site with valid parameters.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC44_UpdateContentTypeWithValidContentType()
        {
            #region Prepare for UpdateContentType.
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);
            this.UpdateFields = this.GenerateUpdateFields();
            this.NewFields = this.GenerateNewFieldsForUpdate();
            UpdateContentTypeContentTypeProperties newProperty = GenerateContentTypeProperties();
            this.DeleteField = this.GenerateDeleteFieldsDefinition();
            #endregion

            // Update a content type on the context site.
            UpdateContentTypeResponseUpdateContentTypeResult updateContentType = Adapter.UpdateContentType(contentTypeID, newProperty, NewFields, UpdateFields, DeleteField);

            #region Capture requirements.

            // Verify MS-WEBSS requirement: MS-WEBSS_R555
            // If the UpdateContentResult is not null, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                updateContentType.Results,
                555,
                @"[In UpdateContentTypeResponse] UpdateContentTypeResult: If no error conditions as specified previously cause the protocol server to return a SOAP exception, an UpdateContentTypeResult MUST be returned.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R556, MS-WEBSS_R558
            bool isVerifiedMethodId = false;
            int len = updateContentType.Results.Method.Length;
            for (int i = 0; i < len; i++)
            {
                if (updateContentType.Results.Method[i].ID.Equals(NewFields.Fields[0].ID, StringComparison.OrdinalIgnoreCase)
                    || updateContentType.Results.Method[i].ID.Equals(UpdateFields.Fields[0].ID, StringComparison.OrdinalIgnoreCase)
                    || updateContentType.Results.Method[i].ID.Equals(DeleteField.Fields[0].ID, StringComparison.OrdinalIgnoreCase))
                {
                    isVerifiedMethodId = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedMethodId,
                556,
                @"[In UpdateContentTypeResponse] Method: Container for data on a field add, update, or remove operation.");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedMethodId,
                558,
                @"[In UpdateContentTypeResponse] Method.ID: The identifier of the Method that is specified in the request.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R529
            Site.CaptureRequirement(
                529,
                @"[In UpdateContentType] The protocol server updates the appropriate fields of the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R530
            Site.CaptureRequirement(
                530,
                @"[In UpdateContentType] The protocol server updates the content type with the specified properties.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R557
            // If the return value ID of method including the keyword "insert", this indicates that the update is successful, 
            // this requirement will be captured.
            bool isVerifyR557 = false;
            for (int i = 0; i < len; i++)
            {
                if (updateContentType.Results.Method[i].ID.Equals(NewFields.Fields[0].ID, StringComparison.OrdinalIgnoreCase))
                {
                    isVerifyR557 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR557,
                557,
                @"[In UpdateContentTypeResponse] Results: Container for data on the update of a content type");

            // Verify MS-WEBSS requirement: MS-WEBSS_R572
            Site.CaptureRequirementIfAreEqual<ONEONLY>(
                ONEONLY.Item1,
                updateContentType.Results.ListProperties.Locs,
                572,
                @"[In UpdateContentTypeResponse] If the content type has localizations, this[ListProperties.Locs] value MUST be set to 1.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R579
            Site.CaptureRequirementIfAreEqual<TRUEONLY>(
                TRUEONLY.TRUE,
                updateContentType.Results.ListProperties.Sealed,
                579,
                @"[In UpdateContentTypeResponse] If this content type is sealed, this[ListProperties.Sealed] value MUST be set to TRUE.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R565
            // If the return value of ListProperties is not null, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                updateContentType.Results.ListProperties,
                565,
                @"[In UpdateContentTypeResponse] ListProperties: Contains updated data for the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R566
            // If the return value of ListProperties is not null, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                updateContentType.Results.ListProperties.Description,
                566,
                @"[In UpdateContentTypeResponse] ListProperties.Description: Description of the content type.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R567
            Site.CaptureRequirementIfAreEqual<bool>(
                true,
                updateContentType.Results.ListProperties.FeatureId == null,
                567,
                @"[In UpdateContentTypeResponse] ListProperties.FeatureId: Specifies the feature with which this content type is associated.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R568
            bool isVerifyR568 = false;
            if (updateContentType.Results.ListProperties.Group.ToString().ToLower().Contains("content type"))
            {
                isVerifyR568 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR568,
                568,
                @"[In UpdateContentTypeResponse] ListProperties.Group: Name of the content type group that contains this[ListProperties.Group] content type.");

            #endregion

            // Update a content type, set newFields,updateFields and deleteFields to null.
            Adapter.UpdateContentType(contentTypeID, newProperty, null, null, null);
        }

        /// <summary>
        /// This test case aims to verify the DeleteContentType operation with valid contentTypeId.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC45_DeleteContentTypeWithValidContentType()
        {
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(DisplayName);

            // Remove a specified content type from the site.
            DeleteContentTypeResponseDeleteContentTypeResult deleteResult = Adapter.DeleteContentType(contentTypeID);

            // Verify MS-WEBSS requirement: MS-WEBSS_R119
            // If the return result is not null, the operation succeeds, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                deleteResult,
                119,
                @"[In DeleteContentTypeResponse] DeleteContentTypeResult: If the operation succeeds, a DeleteContentTypeResult MUST be returned.");
        }

        /// <summary>
        /// This test case aims to verify the optional behaviors of GetContentTypes operation.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S01_TC46_GetContentTypesWithShouldMay()
        {
            // Obtain content type data for a specified content type identifier.
            GetContentTypesResponseGetContentTypesResult getContentTypes = Adapter.GetContentTypes();

            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1029, this.Site), "This operation GetContentTypes failed.");
            if (Common.IsRequirementEnabled(1029, this.Site))
            {
                if (getContentTypes.ContentTypes != null)
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1029
                    // When the System Under Test product name is Windows SharePoint Services 3.0 and above, if the server returns a
                    //  positive information when invoke GetContentType operation, then the requirement can be captured.
                    Site.CaptureRequirement(
                        1029,
                        @"[In Appendix B: Product Behavior] Implementation does support this[GetContentTypes] operation.(<16>Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }
    }
}