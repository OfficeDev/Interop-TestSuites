namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with XML document.
    /// </summary>
    [TestClass]
    public class S02_OperationsOnContentTypeXmlDocument : TestSuiteBase
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
        /// This test case aims to verify the RemoveContentTypeXmlDocument operation with invalid content type ID.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC01_RemoveContentTypeXmlDocumentInvalidWithEmpty()
        {
            try
            {
                // Remove a document from the document collection of a site content type. 
                Adapter.RemoveContentTypeXmlDocument(string.Empty, this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP exception is not returned for the RemoveContentTypeXmlDocument operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R386
                Site.CaptureRequirement(
                    386,
                    @"[In RemoveContentTypeXmlDocument] If the content type specified by the contentTypeId is not found, the protocol server MUST return a SOAP exception.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R402
                Site.CaptureRequirement(
                    402,
                    @"[In RemoveContentTypeXmlDocumentResponse] If the operation fails, a SOAP exception MUST be returned.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentTypeXmlDocument operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC02_UpdateContentTypeXmlDocument_Unauthenticated()
        {
            string contentTypeID = CreateContentType(this.GenerateRandomString(10));
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                XmlDocument rawXmlDoc = new XmlDocument();
                Adapter.UpdateContentTypeXmlDocument(contentTypeID, rawXmlDoc.DocumentElement);
                Site.Assert.Fail("The expected http status code is not returned for the UpdateContentTypeXmlDocument operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1088
                // COMMENT: When the UpdateContentTypeXmlDocument operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1088,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[UpdateContentTypeXmlDocument], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentTypeXmlDocument operation with invalid contentTypeId.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC03_UpdateContentTypeXmlDocumentInvalidContentType()
        {
            string customInfo = @"<customInfo xmlns='http://www.contoso.com/customInfo'>Here is some custom information</customInfo>";
            XmlDocument rawXmlDoc = new XmlDocument();
            rawXmlDoc.LoadXml(customInfo);
            XmlElement rawResponseXml = rawXmlDoc.DocumentElement;

            // Create a new content type on the context site.
            this.CreateContentType(this.GenerateRandomString(10));

            try
            {
                Adapter.UpdateContentTypeXmlDocument(this.GenerateRandomString(6), rawResponseXml);
                Site.Assert.Fail("The expected SOAP exception is not returned for the UpdateContentTypeXmlDocument operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R588
                Site.CaptureRequirement(
                    588,
                   @"[In UpdateContentTypeXmlDocument] If the content type specified by the contentTypeId is not found, the protocol server MUST return a SOAP exception.");

                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R616
                Site.CaptureRequirement(
                    616,
                    @"[In UpdateContentTypeXmlDocumentResponse] If the operation fails, a SOAP exception MUST be returned.");
            }
        }

        /// <summary>
        /// This test case aims to verify the UpdateContentTypeXmlDocument operation with invalid newDocument value.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC04_UpdateContentTypeXmlDocumentInvalidXmlElement()
        {
            string malformedCustomInfo = @"<custom xml='http://www.contoso.com/customInfo'><customInfo></customInfo><Error/>Here is some custom information</custom>";
            XmlDocument rawXmlDoc = new XmlDocument();
            rawXmlDoc.LoadXml(malformedCustomInfo);
            XmlElement rawResponseXml = rawXmlDoc.DocumentElement;

            // Create a new content type on the context site.
            this.CreateContentType(this.GenerateRandomString(10));

            try
            {
                Adapter.UpdateContentTypeXmlDocument(this.GenerateRandomString(6), rawResponseXml);
                Site.Assert.Fail("The expected SOAP exception is not returned for the UpdateContentTypeXmlDocument operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R590
                Site.CaptureRequirement(
                    590,
                    @"[In UpdateContentTypeXmlDocument] If newDocument is malformed XML, the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the RemoveContentTypesXmlDocument operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC05_RemoveContentTypeXmlDocument_Unauthenticated()
        {
            string contentTypeID = CreateContentType(this.GenerateRandomString(10));
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.RemoveContentTypeXmlDocument(contentTypeID, this.GenerateRandomString(10));
                Site.Assert.Fail("The expected http status code is not returned for the RemoveContentTypeXmlDocument operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1082
                // COMMENT: When the RemoveContentTypeXmlDocument operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1082,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[RemoveContentTypeXmlDocument], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify RemoveContentTypeXmlDocument operation to remove a document from the document collection of a site content type.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC06_RemoveContentTypeXmlDocument()
        {
            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(this.GenerateRandomString(10));

            RemoveContentTypeXmlDocumentResponseRemoveContentTypeXmlDocumentResult actual = new RemoveContentTypeXmlDocumentResponseRemoveContentTypeXmlDocumentResult();

            // Remove a document from the document collection of a site content type. 
            actual = Adapter.RemoveContentTypeXmlDocument(contentTypeID, this.GenerateRandomString(10));

            // Verify MS-WEBSS requirement: MS-WEBSS_R389
            Site.CaptureRequirementIfIsNotNull(
                actual.Success,
                389,
                @"[In RemoveContentTypeXmlDocument] If an XML document in the requested content type has the namespace specified by the documentUri, it is removed from the document collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R390
            Site.CaptureRequirementIfIsTrue(
                actual.Success.ToString().Contains("RemoveContentTypeXmlDocumentResult"),
                390,
                @"[In RemoveContentTypeXmlDocument] If no error is raised, the protocol server MUST return a success RemoveContentTypeXmlDocumentResult.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R401
            Site.CaptureRequirementIfIsTrue(
                actual.Success.ToString().Contains("RemoveContentTypeXmlDocumentResult"),
                401,
                @"[In RemoveContentTypeXmlDocumentResponse] RemoveContentTypeXmlDocumentResult: If the operation succeeds, a RemoveContentTypeXmlDocumentResult MUST be returned.");

            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1033, this.Site), "This operation RemoveContentTypeXmlDocument failed.");
            if (Common.IsRequirementEnabled(1033, this.Site))
            {
                // If the operation succeed, Verify MS-WEBSS requirement: MS-WEBSS_R1033
                this.Site.CaptureRequirement(
                    1033,
                    @"[In Appendix B: Product Behavior]  Implementation does support this [RemoveContentTypeXmlDocument] operation.(<19>Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case aims to verify UpdateContentTypeXmlDocument operation to update an XML document in the XML Document collection of a site content type.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S02_TC07_UpdateContentTypeXmlDocument()
        {
            string customInfo = @"<customInfo xmlns='http://www.contoso.com/customInfo'>Here is some custom information</customInfo>";
            XmlDocument rawXmlDoc = new XmlDocument();
            rawXmlDoc.LoadXml(customInfo);
            XmlElement rawResponseXml = rawXmlDoc.DocumentElement;

            // Create a new content type on the context site.
            string contentTypeID = CreateContentType(this.GenerateRandomString(10));
            UpdateContentTypeXmlDocumentResponseUpdateContentTypeXmlDocumentResult result = Adapter.UpdateContentTypeXmlDocument(contentTypeID, rawResponseXml);
            result = Adapter.UpdateContentTypeXmlDocument(contentTypeID, rawResponseXml);

            // Verify MS-WEBSS requirement: MS-WEBSS_R615
            Site.CaptureRequirementIfIsTrue(
                result.Success.ToString().Contains("UpdateContentTypeXmlDocumentResult"),
                615,
                @"[In UpdateContentTypeXmlDocumentResponse] UpdateContentTypeXmlDocumentResult: If the operation succeeds, an UpdateContentTypeXmlDocumentResult element MUST be returned.");

            Site.Assert.IsFalse(!Common.IsRequirementEnabled(1039, this.Site), "This operation UpdateContentTypeXmlDocument failed.");
            if (Common.IsRequirementEnabled(1039, this.Site))
            {
                // If the operation succeed, Verify MS-WEBSS requirement: MS-WEBSS_R1039
                Site.CaptureRequirement(
                    1039,
                    @"[In Appendix B: Product Behavior] Implementation does support this[UpdateContentTypeXmlDocument] operation.(<25>Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }
    }
}