namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with columns. 
    /// </summary>
    [TestClass]
    public class S07_OperationsOnColumns : TestSuiteBase
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
        /// This test case aims to verify the UpdateColumns operation with invalid updateFields, updateFields or deleteFields which specify an invalid FieldDefinition element.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC01_UpdateColumns_InvalidFieldDefinition()
        {
            #region Set up the environment.
            this.InitUpdateColumn();

            UpdateColumnsMethod[] newFields = new UpdateColumnsMethod[1];
            newFields[0] = new UpdateColumnsMethod();
            newFields[0].Field = new FieldDefinition();
            newFields[0].ID = ((uint)CONST_MethodIDUNITS.One).ToString();
            newFields[0].Field.Name = this.GenerateRandomString(10);
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[1];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.FieldNameForUpdate;
            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[1];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(newFields, updateFields, deleteFields);

            #region Capture Invalid Field Definition Related Requirement
            bool isCorrectResponseResult = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.NewFields[0].ErrorCode)
                && (updateColumnsResult.Results.NewFields[0].ErrorText != null)
                && (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.UpdateFields[0].ErrorCode)
                && (updateColumnsResult.Results.UpdateFields[0].ErrorText != null)
                && (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.DeleteFields[0].ErrorCode)
                && (updateColumnsResult.Results.DeleteFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
                SoapErrorCode.ErrorCode0x80004005,
                updateColumnsResult.Results.NewFields[0].ErrorCode,
                "The expected error code {0} should be returned in NewFields of the UpdateColumnsResponse element.",
                SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.NewFields[0].ErrorText, "The ErrorText of NewFields in UpdateColumnsResponse should not be null");

            Site.Assert.AreEqual<string>(
               SoapErrorCode.ErrorCode0x80004005,
               updateColumnsResult.Results.DeleteFields[0].ErrorCode,
               "The expected error code {0} of DeleteFields should be returned in the UpdateColumnsResponse element.",
               SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.DeleteFields[0].ErrorText, "The ErrorText of DeleteFields in UpdateColumnsResponse should not be null");

            Site.Assert.AreEqual<string>(
              SoapErrorCode.ErrorCode0x80004005,
              updateColumnsResult.Results.UpdateFields[0].ErrorCode,
              "The expected error code {0} of UpdateFields should be returned in the UpdateColumnsResponse element.",
              SoapErrorCode.ErrorCode0x80004005);
            Site.Assert.IsNotNull(updateColumnsResult.Results.UpdateFields[0].ErrorText, "The ErrorText of UpdateFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R466
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                466,
                @"[In UpdateColumns]  If an error occurs, the protocol server MUST return an appropriate error code and error string.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R826
            this.Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                826,
                @"[In UpdateColumnsResponse] If the protocol server encounters one of the error conditions[An invalid Field element is passed in any of the parameters.] in the following table while running this operation[UpdateColumns], ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element.");

            if (Common.IsRequirementEnabled(834, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R834
                this.Site.CaptureRequirementIfIsTrue(
                    isCorrectResponseResult,
                    834,
                    @"[In UpdateColumnsResponse] If implementation does encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition[An invalid Field element is passed in any of the parameters.].
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the GetColumns operation when site refer to service has invalid column information.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC02_GetColumns_WithInvalidColumn()
        {
            #region Set up the environment
            this.InitUpdateColumn();

            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[1];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.FieldNameForUpdate;
            Adapter.UpdateColumns(null, updateFields, null);
            #endregion

            try
            {
                Adapter.GetColumns();
                Site.Assert.Fail("The expected SOAP fault is not returned for the GetListTemplates operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R165
                Site.CaptureRequirement(
                    165,
                    @"[In GetColumns] If the operation fails, the protocol server MUST return a SOAP exception.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R166
                Site.CaptureRequirement(
                    166,
                    @"[In GetColumns] A SOAP fault MUST be returned when a GetColumns operation is performed on a context site that has invalid column attribute information.");
            }
        }

        /// <summary>
        /// This test case aims to verify UpdateColumns operation to add, update, or delete one or more specified existing columns on the context site and all child sites within its hierarchy.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC03_UpdateColumns_AllFieldsValid()
        {
            #region Set up the environment
            this.InitUpdateColumn();

            // Update to new interface
            UpdateColumnsMethod[] newFields = new UpdateColumnsMethod[1];
            newFields[0] = new UpdateColumnsMethod();
            newFields[0].Field = new FieldDefinition();
            newFields[0].ID = ((uint)CONST_MethodIDUNITS.One).ToString();
            newFields[0].Field.Name = this.GenerateRandomString(10);
            newFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);

            // Update to new interface
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[1];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.GenerateRandomString(10);
            updateFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            updateFields[0].Field.DisplayName = this.GenerateRandomString(10);

            // Update to new interface
            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[1];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            deleteFields[0].Field.Name = this.GenerateRandomString(10);

            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(newFields, updateFields, deleteFields);

            #region Capture All Fields Valid Related Requirement
            // Verify MS-WEBSS requirement: MS-WEBSS_R465
            bool isVerifiedR465 = updateColumnsResult != null;
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR465,
                465,
                @"[In UpdateColumns] If the operation succeeds, it MUST return an UpdateColumnsResponse element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R508
            Site.CaptureRequirementIfAreEqual<string>(
                newFields[0].ID,
               updateColumnsResult.Results.NewFields[0].ID,
                508,
                @"[In UpdateColumnsResponse] Method.ID: This attribute MUST have the same value as the Method.ID attribute that was sent to the protocol server in the UpdateColumns message of this UpdateColumnsResponse for an add operation.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R473
            Site.CaptureRequirementIfAreEqual<string>(
                newFields[0].ID,
                updateColumnsResult.Results.NewFields[0].ID,
                473,
                @"[In UpdateColumnsSoapOut] This message[UpdateColumnsSoapOut] is the response message to perform the following operations on the context site and all child sites in its hierarchy:
	Adding one or more specified new columns.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R837
            Site.CaptureRequirementIfAreEqual<string>(
                updateFields[0].ID,
                updateColumnsResult.Results.UpdateFields[0].ID,
                837,
                @"[In UpdateColumnsResponse] Method.ID: This attribute MUST have the same value as the Method.ID attribute that was sent to the protocol server in the UpdateColumns message of this UpdateColumnsResponse for an update operation.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R474
            Site.CaptureRequirementIfAreEqual<string>(
                updateFields[0].ID,
                updateColumnsResult.Results.UpdateFields[0].ID,
                474,
                @"[In UpdateColumnsSoapOut] Updating one or more specified existing columns.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R838
            Site.CaptureRequirementIfAreEqual<string>(
                deleteFields[0].ID,
                updateColumnsResult.Results.DeleteFields[0].ID,
               838,
                @"[In UpdateColumnsResponse] Method.ID: This attribute MUST have the same value as the Method.ID attribute that was sent to the protocol server in the UpdateColumns message of this UpdateColumnsResponse for a delete operation.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R475
            Site.CaptureRequirementIfAreEqual<string>(
                deleteFields[0].ID,
                updateColumnsResult.Results.DeleteFields[0].ID,
                475,
                @"[In UpdateColumnsSoapOut] Deleting one or more specified existing columns.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R1037
            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1037, this.Site), "This operation UpdateColumns failed.");
            if (Common.IsRequirementEnabled(1037, this.Site))
            {
                // If the operation UpdateColumns failed,verify MS-WEBSS requirement: MS-WEBSS_R1037
                Site.CaptureRequirement(
                    1037,
                    @"[In Appendix B: Product Behavior] Implementation does support this[UpdateColumns] operation.(<21> Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation when the attribute Name of deleteFields and updateFields doesn’t included in existing columns.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC04_UpdateColumns_NoMatchingName()
        {
            #region Set up the environment.
            this.InitUpdateColumn();
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[1];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.GenerateRandomString(10);
            updateFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);

            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[1];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            deleteFields[0].Field.Name = this.GenerateRandomString(10);
            #endregion.

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(null, updateFields, deleteFields);

            bool iscorrectResponseResultForDelete = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.DeleteFields[0].ErrorCode)
                && (updateColumnsResult.Results.DeleteFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
               SoapErrorCode.ErrorCode0x80004005,
               updateColumnsResult.Results.DeleteFields[0].ErrorCode,
               "The expected error code {0} of DeleteFields should be returned in the UpdateColumnsResponse element.",
               SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.DeleteFields[0].ErrorText, "The ErrorText of DeleteFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R832
            if (Common.IsRequirementEnabled(830, this.Site))
            {
                this.Site.CaptureRequirementIfIsTrue(
                    iscorrectResponseResultForDelete,
                    832,
                    @"[In UpdateColumnsResponse] If implementation does encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition. [An invalid GUID is passed in as the ID attribute for updateFields and deleteFields.]
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetColumns operation for the user without authorization.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC05_GetColumns_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetColumns();
                Site.Assert.Fail("The expected http status code is not returned for the GetColumns operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1074
                // COMMENT: When the GetColumns operation is invoked by unauthenticated user, if the 
                // server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1074,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetColumns], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation for the user without authorization.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC06_UpdateColumns_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                UpdateColumnsMethod[] newFileds = new UpdateColumnsMethod[1];
                UpdateColumnsMethod1[] updateFileds = new UpdateColumnsMethod1[1];
                UpdateColumnsMethod2[] deleteFileds = new UpdateColumnsMethod2[1];
                Adapter.UpdateColumns(
                    newFileds,
                    updateFileds,
                    deleteFileds);
                Site.Assert.Fail("The expected http status code is not returned for the UpdateColumns operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1086
                // COMMENT: When the UpdateColumns operation is invoked by unauthenticated user, if the 
                // server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                     1086,
                 @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[UpdateColumns], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid newFields which doesn’t have a root element.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC07_UpdateColumns_NewFieldsWithMultipleMethodNoFields()
        {
            #region Set up the environment
            this.InitUpdateColumn();
            UpdateColumnsMethod[] newFields = new UpdateColumnsMethod[2];
            newFields[0] = new UpdateColumnsMethod();
            newFields[0].Field = new FieldDefinition();
            newFields[0].Field.Name = this.GenerateRandomString(10);
            newFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            newFields[0].ID = ((uint)CONST_MethodIDUNITS.One).ToString();

            newFields[1] = new UpdateColumnsMethod();
            newFields[1].Field = new FieldDefinition();
            newFields[1].Field.Name = this.GenerateRandomString(10);
            newFields[1].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            newFields[1].ID = ((uint)CONST_MethodIDUNITS.Two).ToString();
            
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[2];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.GenerateRandomString(10);
            updateFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            updateFields[0].Field.DisplayName = this.GenerateRandomString(10);

            updateFields[1] = new UpdateColumnsMethod1();
            updateFields[1].Field = new FieldDefinition();
            updateFields[1].ID = ((uint)CONST_MethodIDUNITS.Four).ToString();
            updateFields[1].Field.Name = this.GenerateRandomString(10);
            updateFields[1].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            updateFields[1].Field.DisplayName = this.GenerateRandomString(10);
            
            #endregion

            try
            {
                Adapter.UpdateColumns(newFields, null, null);
                Site.Assert.Fail("The expected SOAP fault is not returned for the UpdateColumns operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R515
                Site.CaptureRequirement(
                    515,
                    @"[In UpdateColumnsResponse] [A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation] Occurs when one of the newFields  elements of the UpdateColumns element has multiple Method elements without a Fields element defined as the root element.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R514
                Site.CaptureRequirement(
                    514,
                     @"[In UpdateColumnsResponse] When an invalid XML element is passed in as newFields element , a SOAP fault MUST be returned.");
            }

            try
            {
                Adapter.UpdateColumns(null,updateFields, null);
                Site.Assert.Fail("The expected SOAP fault is not returned for the UpdateColumns operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R822
                Site.CaptureRequirement(
                    822,
                    @"[In UpdateColumnsResponse] When an invalid element is the child element of the [newFields, ]updateFields[, or deleteFields elements],  a SOAP fault MUST be returned.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid updateFields which doesn’t have a root element.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC08_UpdateColumns_UpdateFieldWithMultipleMethodNoFields()
        {
            #region Set up the environment
            this.InitUpdateColumn();
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[2];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].Field.Name = this.GenerateRandomString(10);
            updateFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            updateFields[0].Field.DisplayName = this.GenerateRandomString(10);
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();

            updateFields[1] = new UpdateColumnsMethod1();
            updateFields[1].Field = new FieldDefinition();
            updateFields[1].Field.Name = this.GenerateRandomString(10);
            updateFields[1].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            updateFields[1].Field.DisplayName = this.GenerateRandomString(10);
            updateFields[1].ID = ((uint)CONST_MethodIDUNITS.Four).ToString();

            #endregion

            try
            {
                Adapter.UpdateColumns(null, updateFields, null);
                Site.Assert.Fail("The expected SOAP fault is not returned for the UpdateColumns operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R734
                Site.CaptureRequirement(
                    734,
                   @"[In UpdateColumnsResponse] [A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation] Occurs when one of the updateFields  elements of the UpdateColumns element has multiple Method elements without a Fields element defined as the root element.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R1090
                Site.CaptureRequirement(
                    1090,
                @"[In UpdateColumnsResponse] [A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation] Occurs when one of the updateFields elements of the UpdateColumns element has multiple Method elements without a Fields element defined as the root element.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid DeleteFields which doesn’t have a root element.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC09_UpdateColumns_DeleteFieldsWithMultipleMethodNoFields()
        {
            #region Set up the environment
            this.InitUpdateColumn();
            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[2];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].Field.Name = this.GenerateRandomString(10);
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();

            deleteFields[1] = new UpdateColumnsMethod2();
            deleteFields[1].Field = new FieldDefinition();
            deleteFields[1].Field.Name = this.GenerateRandomString(10);
            deleteFields[1].ID = ((uint)CONST_MethodIDUNITS.Six).ToString();

            #endregion

            try
            {
                Adapter.UpdateColumns(null, null, deleteFields);
                Site.Assert.Fail("The expected SOAP fault is not returned for the UpdateColumns operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R735
                Site.CaptureRequirement(
                    735,
                    @"[In UpdateColumnsResponse] [A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation] Occurs when one of the deleteFields elements of the UpdateColumns element has multiple Method elements without a Fields element defined as the root element.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R822001
                Site.CaptureRequirement(
                    822001,
                    @"[In UpdateColumnsResponse] When an invalid element is the child element of the [newFields, updateFields, or ]deleteFields elements,  a SOAP fault MUST be returned.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid newFields which specifies an already existing column.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC10_UpdateColumns_ExistingColumn()
        {
            #region Set up the environment.
            this.InitUpdateColumn();
            UpdateColumnsMethod[] newFields = new UpdateColumnsMethod[1];
            newFields[0] = new UpdateColumnsMethod();
            newFields[0].Field = new FieldDefinition();
            newFields[0].ID = ((uint)CONST_MethodIDUNITS.One).ToString();
            newFields[0].Field.Name = this.GenerateRandomString(10);
            newFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            Adapter.UpdateColumns(newFields, null, null);
            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(newFields, null, null);

            #region Capture Add Existing Column Related Requirement
            bool isCorrectResponseResult = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.NewFields[0].ErrorCode) && (updateColumnsResult.Results.NewFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
                SoapErrorCode.ErrorCode0x80004005,
                updateColumnsResult.Results.NewFields[0].ErrorCode,
                "The expected error code {0} should be returned in NewFields of the UpdateColumnsResponse element.",
                SoapErrorCode.ErrorCode0x80004005);
            Site.Assert.IsNotNull(updateColumnsResult.Results.NewFields[0].ErrorText, "The ErrorText of NewFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R466
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                466,
                @"[In UpdateColumns]  If an error occurs, the protocol server MUST return an appropriate error code and error string.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R516
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                516,
                @"[In UpdateColumnsResponse] If the protocol server encounters one of the error conditions[An attempt is made to add an already existing column to the site.] in the following table while running this operation[UpdateColumns], ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element.");

            if (Common.IsRequirementEnabled(830, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R830
                this.Site.CaptureRequirementIfIsTrue(
                    isCorrectResponseResult,
                    830,
                    @"[In UpdateColumnsResponse] If implementation does  encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition. [An attempt is made to add an already existing column to the site.]
[ The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid deleteFields which specifies an already existing column.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC11_UpdateColumns_NonexistentColumn()
        {
            #region Set up the environment.
            this.InitUpdateColumn();
            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[1];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            deleteFields[0].Field.ID = Guid.NewGuid().ToString();

            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(null, null, deleteFields);

            #region Capture Delete a Nonexistent Column Related Requirement
            bool isCorrectResponseResult = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.DeleteFields[0].ErrorCode)
                && (updateColumnsResult.Results.DeleteFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
               SoapErrorCode.ErrorCode0x80004005,
               updateColumnsResult.Results.DeleteFields[0].ErrorCode,
               "The expected error code {0} of DeleteFields should be returned in the UpdateColumnsResponse element.",
               SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.DeleteFields[0].ErrorText, "The ErrorText of DeleteFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R466
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                466,
                @"[In UpdateColumns]  If an error occurs, the protocol server MUST return an appropriate error code and error string.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R823
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                823,
                 @"[In UpdateColumnsResponse] If the protocol server encounters one of the error conditions[An attempt is made to delete a non-existing column from the site.] in the following table while running this operation[UpdateColumns], ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element.");

            if (Common.IsRequirementEnabled(831, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R831
                this.Site.CaptureRequirementIfIsTrue(
                    isCorrectResponseResult,
                    831,
                    @"[In UpdateColumnsResponse] If implementation does  encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition. [An attempt is made to delete a non-existing column from the site.]
 [The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation with invalid GUID in deleteFields and updateFields.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC12_UpdateColumns_InvalidGUID()
        {
            #region Set up the environment.
            this.InitUpdateColumn();
            UpdateColumnsMethod1[] updateFields = new UpdateColumnsMethod1[1];
            updateFields[0] = new UpdateColumnsMethod1();
            updateFields[0].Field = new FieldDefinition();
            updateFields[0].ID = ((uint)CONST_MethodIDUNITS.Three).ToString();
            updateFields[0].Field.Name = this.GenerateRandomString(10);
            updateFields[0].Field.ID = Guid.NewGuid().ToString();
            UpdateColumnsMethod2[] deleteFields = new UpdateColumnsMethod2[1];
            deleteFields[0] = new UpdateColumnsMethod2();
            deleteFields[0].Field = new FieldDefinition();
            deleteFields[0].ID = ((uint)CONST_MethodIDUNITS.Five).ToString();
            deleteFields[0].Field.ID = Guid.NewGuid().ToString();

            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(null, updateFields, deleteFields);

            #region Capture Invalid GUID Related Requirement
            bool isCorrectResponseResult = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.DeleteFields[0].ErrorCode)
                && (updateColumnsResult.Results.DeleteFields[0].ErrorText != null)
                && (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.UpdateFields[0].ErrorCode)
                && (updateColumnsResult.Results.UpdateFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
               SoapErrorCode.ErrorCode0x80004005,
               updateColumnsResult.Results.DeleteFields[0].ErrorCode,
               "The expected error code {0} of DeleteFields should be returned in the UpdateColumnsResponse element.",
               SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.DeleteFields[0].ErrorText, "The ErrorText of DeleteFields in UpdateColumnsResponse should not be null");

            Site.Assert.AreEqual<string>(
               SoapErrorCode.ErrorCode0x80004005,
               updateColumnsResult.Results.UpdateFields[0].ErrorCode,
               "The expected error code {0} of UpdateFields should be returned in the UpdateColumnsResponse element.",
               SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.UpdateFields[0].ErrorText, "The ErrorText of UpdateFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R466
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                466,
                @"[In UpdateColumns]  If an error occurs, the protocol server MUST return an appropriate error code and error string.");

            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                467,
                @"[In UpdateColumns] Error code(s) specific to this operation[UpdateColumns] are defined in UpdateColumnsResponse (section 3.1.4.18.2.2).");

            // Verify MS-WEBSS requirement: MS-WEBSS_R824
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                824,
                @"[In UpdateColumnsResponse] If the protocol server encounters one of the error conditions[An invalid GUID is passed in as the ID attribute for updateFields and deleteFields.] in the following table while running this operation[UpdateColumns], ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element.");

            if (Common.IsRequirementEnabled(832, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R832
                this.Site.CaptureRequirementIfIsTrue(
                    isCorrectResponseResult,
                    832,
                    @"[In UpdateColumnsResponse] If implementation does encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition. [An invalid GUID is passed in as the ID attribute for updateFields and deleteFields.]
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }
            #endregion
        }

        /// <summary>
        /// This test case aims to verify the UpdateColumns operation when the attributes name or display name are not included in newFields.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S07_TC13_UpdateColumns_NoNameOrDisplayName()
        {
            #region Set up the environment.
            UpdateColumnsMethod[] newFields = new UpdateColumnsMethod[1];
            newFields[0] = new UpdateColumnsMethod();
            newFields[0].Field = new FieldDefinition();
            newFields[0].ID = ((uint)CONST_MethodIDUNITS.One).ToString();
            newFields[0].Field.Type = Common.GetConfigurationPropertyValue("UpdateColumns_Type", this.Site);
            #endregion

            UpdateColumnsResponseUpdateColumnsResult updateColumnsResult = Adapter.UpdateColumns(newFields, null, null);

            #region Capture No Name Or Display Name of new Fields Related Requirement

            bool isCorrectResponseResult = (SoapErrorCode.ErrorCode0x80004005 == updateColumnsResult.Results.NewFields[0].ErrorCode)
                && (updateColumnsResult.Results.NewFields[0].ErrorText != null);

            Site.Assert.AreEqual<string>(
                SoapErrorCode.ErrorCode0x80004005,
                updateColumnsResult.Results.NewFields[0].ErrorCode,
                "The expected error code {0} should be returned in NewFields of the UpdateColumnsResponse element.",
                SoapErrorCode.ErrorCode0x80004005);

            Site.Assert.IsNotNull(updateColumnsResult.Results.NewFields[0].ErrorText, "The ErrorText of NewFields in UpdateColumnsResponse should not be null");

            // Verify MS-WEBSS requirement: MS-WEBSS_R466
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                466,
                @"[In UpdateColumns]  If an error occurs, the protocol server MUST return an appropriate error code and error string.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R825
            Site.CaptureRequirementIfIsTrue(
                isCorrectResponseResult,
                825,
                @"[In UpdateColumnsResponse] If the protocol server encounters one of the error conditions[Neither the Name nor the DisplayName attribute is passed in a Field element of newFields.] in the following table while running this operation[UpdateColumns], ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element.");

            if (Common.IsRequirementEnabled(833, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R833
                this.Site.CaptureRequirementIfIsTrue(
                    isCorrectResponseResult,
                    833,
                    @"[In UpdateColumnsResponse] If implementation does encounter one of the error conditions in the following table while running this operation, ErrorCode[0x80004005] and ErrorText elements MUST be returned in the UpdateColumnsResponse element, which contain one of the error codes in the following table for the specified error condition[Neither the Name nor the DisplayName attribute is passed in a Field element of newFields.].
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
            }

            #endregion
        }
    }
}