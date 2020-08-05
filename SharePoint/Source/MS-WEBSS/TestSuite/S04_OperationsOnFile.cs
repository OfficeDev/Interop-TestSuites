namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with file.
    /// </summary>
    [TestClass]
    public class S04_OperationsOnFile : TestSuiteBase
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
        /// This test case aims to verify RevertAllFileContentStreams operation to let all pages within the context site revert to their original states.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC01_RevertAllFileContentStreams()
        {
            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1034, this.Site), "This operation RevertAllFileContentStreams failed.");
            if (Common.IsRequirementEnabled(1034, this.Site))
            {
                // For capture requirements in adapter.
                Adapter.RevertAllFileContentStreams();

                // If the product is Windows® SharePoint® Services 3.0 and above, Verify MS-WEBSS requirement: MS-WEBSS_R1034
                Site.CaptureRequirement(
                    1034,
                    @"[In Appendix B: Product Behavior]  Implementation does support this[RevertAllFileContentStreams] operation.(<20>Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case aims to verify RevertFileContentStream operation to let the specified page within the context site revert to its original state.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC02_RevertFileContentStream_ValidFileUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1036, this.Site), "The test case is executed only when the property 'R1036Enabled' is true.");

            // For capture requirements in adapter.
            Adapter.RevertFileContentStream(Common.GetConfigurationPropertyValue("RevertFileContentStream_ValidFileUrl", this.Site));

            // If the product is Windows® SharePoint® Services 3.0 and above, Verify MS-WEBSS requirement: MS-WEBSS_R1036 when invokes the operation "RevertFileContentStream" successfully.
            Site.CaptureRequirement(
                 1036,
                @"[In Appendix B: Product Behavior] Implementation does support this[RevertFileContentStream] operation.(<22>Windows SharePoint Services 3.0 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with invalid parameter.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC03_RevertFileContentStream_InvalidPageUrl()
        {
            try
            {
                Adapter.RevertFileContentStream(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R726
                // Add the debug information
                Site.CaptureRequirement(
                    726,
                @"[In RevertFileContentStreamResponse] A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation[RevertFileContentStream]:
 Occurs when an invalid URL for the page is passed in.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with invalid parameter.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC04_RevertFileContentStream_InvalidSiteUrl()
        {
            try
            {
                Adapter.RevertFileContentStream(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R453
                Site.CaptureRequirement(
                   453,
                @"[In RevertFileContentStreamResponse] A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation:
] Occurs when the site referred by the fileUrl is not a valid site.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with valid fileUrl but refers to a page on the parent site of the context site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC05_RevertFileContentStream_ValidUrlReferToParentSite()
        {
            try
            {
                Adapter.RevertFileContentStream(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R454
                Site.CaptureRequirement(
                    454,
                @"[In RevertFileContentStreamResponse]A SOAP fault MUST be returned if the protocol server encounters the following error condition while running this operation:
  Occurs when a valid URL is passed in that refers to a page on the parent site of the context site.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with valid fileUrl which is an empty URL.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC06_RevertFileContentStream_Empty()
        {
            try
            {
                Adapter.RevertFileContentStream(string.Empty);
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException exp)
            {
                Site.Assert.IsNotNull(exp.Detail, "The XmlNode cannot be NULL");

                bool isErrorCodeExist = AdapterHelper.ElementExists((XmlElement)exp.Detail, ConstString.ErrorCode);
                Site.Assert.IsTrue(isErrorCodeExist, "The errorcode element should exist in the SOAPFaultDetails complex type element.");
                bool isErrorStringExist = AdapterHelper.ElementExists((XmlElement)exp.Detail, ConstString.ErrorString);
                Site.Assert.IsTrue(isErrorStringExist, "The errorstring element should exist in the SOAPFaultDetails complex type element.");

                // The error code and error string elements exist, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R439
                Site.CaptureRequirementIfIsTrue(
                    isErrorCodeExist && isErrorStringExist,
                    439,
                @"[In RevertFileContentStream] If there is any problem in performing the operation, the protocol server MUST return appropriate error code and error string.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R440
                Site.CaptureRequirementIfIsTrue(
                    isErrorCodeExist && isErrorStringExist,
                    440,
                    @"[In RevertFileContentStream] Error code(s) specific to this operation[RevertFileContentStream] are defined in the RevertFileContentStreamSoapOut message.");

                if (Common.IsRequirementEnabled(1065, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R813
                    Site.CaptureRequirementIfAreEqual<string>(
                        SoapErrorCode.ErrorCode0x82000001,
                        AdapterHelper.GetSoapExceptionErrorcode(exp),
                        1065,
                    @"[In RevertFileContentStreamResponse] If the protocol server encounters the  error condition: Occurs when an empty URL is passed in or no fileUrl element is supplied, while running this operation[RevertFileContentStream], a SOAP fault MUST be returned that contains the error code 0x82000001 in the following table for the specified error condition.");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with invalid fileUrl which is NULL.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC07_RevertFileContentStream_NULL()
        {
            try
            {
                Adapter.RevertFileContentStream(null);
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException soapExp)
            {
                Site.Assert.IsNotNull(soapExp.Detail, "The XmlNode cannot be NULL");

                bool isErrorCodeExist = AdapterHelper.ElementExists((XmlElement)soapExp.Detail, ConstString.ErrorCode);
                Site.Assert.IsTrue(isErrorCodeExist, "The errorcode element should exist in the SOAPFaultDetails complex type element.");
                bool isErrorStringExist = AdapterHelper.ElementExists((XmlElement)soapExp.Detail, ConstString.ErrorString);
                Site.Assert.IsTrue(isErrorStringExist, "The errorstring element should exist in the SOAPFaultDetails complex type element.");

                // The error code and error string exist, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R439
                Site.CaptureRequirementIfIsTrue(
                    isErrorCodeExist && isErrorStringExist,
                    439,
                @"[In RevertFileContentStream] If there is any problem in performing the operation, the protocol server MUST return appropriate error code and error string.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with valid fileUrl which is not part of site definition.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC08_RevertFileContentStream_NotPartOfSiteDefinition()
        {
            #region Set up the environment
            string foldName = Common.GetConfigurationPropertyValue("FoldName", this.Site);
            string docName = Common.GetConfigurationPropertyValue("DocName", this.Site);
            #endregion

            try
            {
                Adapter.RevertFileContentStream(string.Format("{0}/{1}/{2}", Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site), foldName, docName));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException exp)
            {
                #region Capture SOAP Exception Related Requirement

                Site.Assert.IsNotNull(exp.Detail, "The XmlNode cannot be NULL");

                if (Common.IsRequirementEnabled(1066, this.Site))
                {
                    // Catch the exception, then the following requirements will be captured.
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1066
                    Site.CaptureRequirementIfAreEqual<string>(
                        SoapErrorCode.ErrorCode0x80131600,
                        AdapterHelper.GetSoapExceptionErrorcode(exp),
                        1066,
                    @"[In RevertFileContentStreamResponse]  If the protocol server encounters the error condition: Occurs when the page to be converted is not part of the site definition, while running this operation[RevertFileContentStream], a SOAP fault MUST be returned that contains the error code 0x80131600 in the following table for the specified error condition.");
                }

                bool isErrorCodeExist = AdapterHelper.ElementExists((XmlElement)exp.Detail, ConstString.ErrorCode);
                bool isErrorStringExist = AdapterHelper.ElementExists((XmlElement)exp.Detail, ConstString.ErrorString);
                Site.Assert.IsTrue(isErrorCodeExist || isErrorStringExist, "The errorcode element or errorstring element should exist in the SOAPFaultDetails complex type element.");

                // The error code and error string elements exist, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R439
                Site.CaptureRequirementIfIsTrue(
                    isErrorCodeExist || isErrorStringExist,
                    439,
                @"[In RevertFileContentStream] If there is any problem in performing the operation, the protocol server MUST return appropriate error code and error string.");

                #endregion
            }
        }

        /// <summary>
        /// This test case aims to verify the RevertFileContentStream operation with valid fileUrl which is not referring to page on the content site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC09_RevertFileContentStream_NotReferToPage()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1064, this.Site), "The test case is executed only when the property 'R1064Enabled' is true.");

            try
            {
                Adapter.RevertFileContentStream(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (SoapException e)
            {
                Site.Assert.IsNotNull(e.Detail, "The XmlNode cannot be NULL");

                // Verify MS-WEBSS requirement: MS-WEBSS_R1064
                Site.CaptureRequirementIfIsTrue(
                    e.Detail.LastChild.InnerText.Contains(SoapErrorCode.ErrorCode0x80070002),
                    1064,
                    @"[In RevertFileContentStreamResponse] [In RevertFileContentStreamResponse] If the protocol server encounters the error condition: Occurs when a valid URL is passed in that does NOT refer to a page on the context site, while running this operation[RevertFileContentStream], a SOAP fault MUST be returned that contains the error code 0x80070002 in the following table for the specified error condition.");
            }
        }

        /// <summary>
        /// This test case aims to verify RevertAllFileContentStreams when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC10_RevertAllFileContentStreams_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.RevertAllFileContentStreams();
                Site.Assert.Fail("The expected HTTP status code 401 is not returned for the RevertAllFileContentStreams operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1083
                // COMMENT: When the RevertAllFileContentStreams operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1083,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[RevertAllFileContentStreams], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case is used to test RevertFileContentStream operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S04_TC11_RevertFileContentStream_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.RevertFileContentStream(Common.GetConfigurationPropertyValue("RevertFileContentStream_ValidFileUrl", this.Site));
                Site.Assert.Fail("The expected SOAP fault is not returned for the RevertFileContentStream operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1085
                // COMMENT: When the RevertFileContentStream operation is invoked by unauthenticated user, 
                // if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1085,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[RevertFileContentStream], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }
    }
}