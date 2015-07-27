//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        /// <summary>
        /// An instance of IMS_WEBSSAdapter class which is used to invoke MS-WEBSS operations.
        /// </summary>
        private static IMS_WEBSSAdapter adapter;

        /// <summary>
        /// An instance of IMS_WEBSSSUTControlAdapter class.
        /// </summary>
        private static IMS_WEBSSSUTControlAdapter sutAdapter;

        #region CONSTANTS
        /// <summary>
        /// The identifier of a string.
        /// </summary>
        private string
            contentTypeDescription,
            contentTypeTypeTitle,
            newFieldsID,
            displayName,
            fieldNameForUpdate,
            newFieldName,
            newFieldID;
        #endregion

        /// <summary>
        /// The instance of CreateContentTypeContentTypeProperties.
        /// </summary>
        private CreateContentTypeContentTypeProperties contentTypeType = new CreateContentTypeContentTypeProperties();

        /// <summary>
        /// Gets an instance of IMS_WEBSSAdapter class which is used to invoke MS-WEBSS operations.
        /// </summary>
        protected static IMS_WEBSSAdapter Adapter
        {
            get
            {
                return adapter;
            }
        }

        /// <summary>
        /// Gets an instance of IMS_WEBSSSUTControlAdapter class.
        /// </summary>
        protected static IMS_WEBSSSUTControlAdapter SutAdapter
        {
            get
            {
                return sutAdapter;
            }
        }

        /// <summary>
        /// Gets content type description.
        /// </summary>
        protected string FieldNameForUpdate
        {
            get
            {
                return this.fieldNameForUpdate;
            }
        }

        /// <summary>
        /// Gets content type description.
        /// </summary>
        protected string DisplayName
        {
            get
            {
                return this.displayName;
            }
        }

        /// <summary>
        /// Gets content type description.
        /// </summary>
        protected string NewFieldsID
        {
            get
            {
                return this.newFieldsID;
            }
        }

        /// <summary>
        /// Gets content type description.
        /// </summary>
        protected string ContentTypeTypeTitle
        {
            get
            {
                return this.contentTypeTypeTitle;
            }
        }

        /// <summary>
        /// Gets content type description.
        /// </summary>
        protected string ContentTypeDescription
        {
            get
            {
                return this.contentTypeDescription;
            }
        }

        /// <summary>
        /// Gets or sets container for a list of existing fields to be included in the content type.
        /// </summary>
        protected AddOrUpdateFieldsDefinition NewFields
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets container for a list of existing fields to be included in the content type.
        /// </summary>
        protected AddOrUpdateFieldsDefinition UpdateFields
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets container for a list of fields to be deleted on the content type.
        /// </summary>
        protected DeleteFieldsDefinition DeleteField
        {
            get;
            set;
        }

        /// <summary>
        /// Gets container for properties to set on the content type.
        /// </summary>
        protected CreateContentTypeContentTypeProperties ContentTypeType
        {
            get
            {
                return this.contentTypeType;
            }
        }

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);

            adapter = BaseTestSite.GetAdapter<IMS_WEBSSAdapter>();
            sutAdapter = BaseTestSite.GetAdapter<IMS_WEBSSSUTControlAdapter>();
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        /// <summary>
        /// Initialize the test.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            Common.CheckCommonProperties(this.Site, true);

            Adapter.InitializeService(UserAuthentication.Authenticated);
            this.InitPrivateVariables();
        }

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        /// <summary>
        /// Initialize variables using for test case class.
        /// </summary>
        protected void InitPrivateVariables()
        {
            // Initialize Content type display name.
            this.displayName = Common.GenerateResourceName(this.Site, "ContentType");

            // Initialize Update content type parameters.
            this.contentTypeDescription = this.GenerateRandomString(10);
            this.contentTypeTypeTitle = Common.GenerateResourceName(this.Site, "Title");

            // Initialize AddOrUpdateFieldsDefinition parameters.
            this.newFieldsID = Guid.NewGuid().ToString();
            this.newFieldID = Guid.NewGuid().ToString();
            this.newFieldName = this.GenerateRandomString(10);

            this.fieldNameForUpdate = this.GenerateRandomString(10);
        }

        /// <summary>
        /// Extract error code from error string.
        /// </summary>
        /// <param name="errorString">An error string from soap exception.</param>
        /// <returns>If error string contains error code, error code will be returned; otherwise, empty will be returned.</returns>
        protected string GetErrorCode(string errorString)
        {
            Regex regex = new Regex(@"0x[0-9]*");
            Match m = regex.Match(errorString);
            if (m.Success)
            {
                return m.Value;
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Generate invalid display name for operation CreateContentType.
        /// </summary>
        /// <param name="invalidString">InValid sign.</param>
        /// <returns>Return a display name with special sign.</returns>
        protected string GenerateInvalidDisplayName(CONST_CHARS invalidString)
        {
            return string.Format("{0}{1}{2}", this.GenerateRandomString(3), (char)invalidString, this.GenerateRandomString(3));
        }

        /// <summary>
        /// Generate contentType properties.
        /// </summary>
        /// <returns>Generated content type properties</returns>
        protected UpdateContentTypeContentTypeProperties GenerateContentTypeProperties()
        {
            UpdateContentTypeContentTypeProperties contentTypeProps = new UpdateContentTypeContentTypeProperties();
            contentTypeProps.ContentType = new ContentTypePropertyDefinition();
            contentTypeProps.ContentType.Description = this.GenerateRandomString(10);
            contentTypeProps.ContentType.Title = this.contentTypeTypeTitle;
            return contentTypeProps;
        }

        /// <summary>
        /// Get a list of fields to be deleted on the content type.
        /// </summary>
        /// <returns>The value of deleteField parameter.</returns>
        protected DeleteFieldsDefinition GenerateDeleteFieldsDefinition()
        {
            DeleteFieldsDefinition deleteFields = new DeleteFieldsDefinition();
            deleteFields.Fields = new DeleteFieldsDefinitionMethod[1];
            deleteFields.Fields[0] = new DeleteFieldsDefinitionMethod();
            deleteFields.Fields[0].Field = new DeleteFieldDefinition();
            deleteFields.Fields[0].ID = Guid.NewGuid().ToString();
            deleteFields.Fields[0].Field.Name = this.GenerateRandomString(10);

            return deleteFields;
        }

        /// <summary>
        /// Generate newField parameter with specified newFieldOption.
        /// </summary>
        /// <returns>The value of newField parameter.</returns>
        protected AddOrUpdateFieldsDefinition GenerateNewFields()
        {
            AddOrUpdateFieldsDefinition addOrUpdateFieldsDefinition = new AddOrUpdateFieldsDefinition();
            addOrUpdateFieldsDefinition.Fields = new AddOrUpdateFieldsDefinitionMethod[1];
            addOrUpdateFieldsDefinition.Fields[0] = new AddOrUpdateFieldsDefinitionMethod();
            addOrUpdateFieldsDefinition.Fields[0].ID = this.newFieldsID;

            addOrUpdateFieldsDefinition.Fields[0].Field = new AddOrUpdateFieldDefinition();
            addOrUpdateFieldsDefinition.Fields[0].Field.ID = this.newFieldID;
            addOrUpdateFieldsDefinition.Fields[0].Field.Name = this.newFieldName;

            return addOrUpdateFieldsDefinition;
        }

        /// <summary>
        /// Generate newField parameter with specified newFieldOption.
        /// </summary>
        /// <returns>The value of newField for update operation.</returns>
        protected AddOrUpdateFieldsDefinition GenerateNewFieldsForUpdate()
        {
            AddOrUpdateFieldsDefinition addOrUpdateFieldsDefinition = new AddOrUpdateFieldsDefinition();
            addOrUpdateFieldsDefinition.Fields = new AddOrUpdateFieldsDefinitionMethod[1];
            addOrUpdateFieldsDefinition.Fields[0] = new AddOrUpdateFieldsDefinitionMethod();
            addOrUpdateFieldsDefinition.Fields[0].ID = Guid.NewGuid().ToString();
            addOrUpdateFieldsDefinition.Fields[0].Field = new AddOrUpdateFieldDefinition();
            addOrUpdateFieldsDefinition.Fields[0].Field.ID = this.newFieldID;
            addOrUpdateFieldsDefinition.Fields[0].Field.Name = this.GenerateRandomString(10);

            return addOrUpdateFieldsDefinition;
        }

        /// <summary>
        /// Generate updateFields parameter with specified newFieldOption.
        /// </summary>
        /// <returns>The value of updateFields parameter.</returns>
        protected AddOrUpdateFieldsDefinition GenerateUpdateFields()
        {
            AddOrUpdateFieldsDefinition addOrUpdateFieldsDefinition = new AddOrUpdateFieldsDefinition();
            addOrUpdateFieldsDefinition.Fields = new AddOrUpdateFieldsDefinitionMethod[1];
            addOrUpdateFieldsDefinition.Fields[0] = new AddOrUpdateFieldsDefinitionMethod();
            addOrUpdateFieldsDefinition.Fields[0].ID = this.newFieldsID;

            addOrUpdateFieldsDefinition.Fields[0].Field = new AddOrUpdateFieldDefinition();
            addOrUpdateFieldsDefinition.Fields[0].Field.ID = this.newFieldID;
            addOrUpdateFieldsDefinition.Fields[0].Field.Name = this.GenerateRandomString(10);

            return addOrUpdateFieldsDefinition;
        }

        /// <summary>
        /// Create a ContentType.
        /// </summary>
        /// <param name="contentTypeDisplayName">Display name will store in server.</param>
        /// <returns>The Id of created content type.</returns>
        protected string CreateContentType(string contentTypeDisplayName)
        {
            this.ContentTypeType.ContentType = new ContentTypePropertyDefinition();
            this.ContentTypeType.ContentType.Description = this.contentTypeDescription;
            this.ContentTypeType.ContentType.Title = this.contentTypeTypeTitle;

            AddOrUpdateFieldsDefinition fields = this.GenerateNewFields();

            // Create a new content type on the context site.
            string contentTypeId = Adapter.CreateContentType(contentTypeDisplayName, Common.GetConfigurationPropertyValue("CreateContentType_ValidateContentType", this.Site), fields, this.ContentTypeType);

            return contentTypeId;
        }

        /// <summary>
        /// Do initialization for Updated columns with new fields.
        /// </summary>
        protected void InitUpdateColumn()
        {
            UpdateColumnsMethod[] newFieldsForUpdate = new UpdateColumnsMethod[1];
            newFieldsForUpdate[0] = new UpdateColumnsMethod();
            newFieldsForUpdate[0].Field = new FieldDefinition();
            newFieldsForUpdate[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();

            newFieldsForUpdate[0].Field.Name = this.fieldNameForUpdate;
            newFieldsForUpdate[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            newFieldsForUpdate[0].Field.DisplayName = this.GenerateRandomString(10);
            Adapter.UpdateColumns(newFieldsForUpdate, null, null);

            // Add a column that is used to be deleted in test case.
            UpdateColumnsMethod[] newFieldsForDelete = new UpdateColumnsMethod[1];
            newFieldsForDelete[0] = new UpdateColumnsMethod();
            newFieldsForDelete[0].Field = new FieldDefinition();
            newFieldsForDelete[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            newFieldsForDelete[0].Field.Name = this.GenerateRandomString(10);
            newFieldsForDelete[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            newFieldsForDelete[0].Field.DisplayName = this.GenerateRandomString(10);
            Adapter.UpdateColumns(newFieldsForDelete, null, null);
        }

        /// <summary>
        /// Get Soap version.
        /// </summary>
        /// <returns>The current Soap version</returns>
        protected SoapProtocolVersion GetSoapVersion()
        {
            string soapVersionString = Common.GetConfigurationPropertyValue("SoapVersion", this.Site);
            SoapProtocolVersion soapVersionCurrent = SoapProtocolVersion.Soap12;
            if (string.Compare(soapVersionString, SoapVersion.SOAP11.ToString(), true) == 0)
            {
                soapVersionCurrent = SoapProtocolVersion.Soap11;
            }
            else if (string.Compare(soapVersionString, SoapVersion.SOAP12.ToString(), true) == 0)
            {
                soapVersionCurrent = SoapProtocolVersion.Soap12;
            }

            return soapVersionCurrent;
        }

        /// <summary>
        /// Get invalid URL.
        /// </summary>
        /// <param name="relatedInvalidUrl">Invalid related URL</param>
        /// <returns>Invalid URL</returns>
        protected string GenerateInvalidUrl(string relatedInvalidUrl)
        {
            return Common.GetConfigurationPropertyValue("SiteCollectionUrl", this.Site).ToLower() + ((char)CONST_CHARS.Slash).ToString() + relatedInvalidUrl;
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        protected string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }
    }
}