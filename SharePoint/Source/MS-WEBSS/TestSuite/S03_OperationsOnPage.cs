namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS.Protocol client tries to perform operations associated with page.
    /// </summary>
    [TestClass]
    public class S03_OperationsOnPage : TestSuiteBase
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
        /// This test case aims to verify GetCustomizedPageStatus operation (also known as the ghosted status) of the specified page or file.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S03_TC01_GetCustomizedPageStatus_ValidFileUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1030, this.Site), "The test case is executed only when the property 'R1030Enabled' is true.");

            // For capture requirements in adapter.
            CustomizedPageStatus uncustomizedPageStatus = Adapter.GetCustomizedPageStatus(Common.GetConfigurationPropertyValue("GetCustomizedPageStatus_ValidFileUrl", this.Site));

            // If the product is Windows® SharePoint® Services 3.0 and above, Verify MS-WEBSS requirement: MS-WEBSS_R1030 when invokes the operation "GetCustomizedPageStatus" successfully.
            Site.CaptureRequirement(
                 1030,
                @"[In Appendix B: Product Behavior] Implementation does support this[GetCustomizedPageStatus] operation.(<17> Windows SharePoint Services 3.0 and above follow this behavior.)");

            // Verify MS-WEBSS requirement: MS-WEBSS_R929
            Site.CaptureRequirementIfAreEqual<CustomizedPageStatus>(
                CustomizedPageStatus.Uncustomized,
                uncustomizedPageStatus,
                929,
                @"[In CustomizedPageStatus] [Value] Uncustomized [Meaning] The page specified by fileUrl is not customized on this site.");

            string foldName = Common.GetConfigurationPropertyValue("FoldName", this.Site);
            string docName = Common.GetConfigurationPropertyValue("DocName", this.Site);

            string page = string.Format("{0}/{1}/{2}", Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site), foldName, docName);
            CustomizedPageStatus status = Adapter.GetCustomizedPageStatus(page);

            // Verify MS-WEBSS requirement: MS-WEBSS_R932
            Site.CaptureRequirementIfAreEqual<CustomizedPageStatus>(
                CustomizedPageStatus.None,
                status,
                932,
                @"[In CustomizedPageStatus] [Value] None[Meaning] The specified fileUrl represents a piece of static content, such as an individual text file or a static HTML file.");
        }

        /// <summary>
        /// This test case aims to verify the GetCustomizedPageStatus operation with invalid parameter.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S03_TC02_GetCustomizedPageStatus_InvalidFileUrl()
        {
            try
            {
                Adapter.GetCustomizedPageStatus(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the GetCustomizedPageStatus operation.");
            }
            catch (SoapException)
            {
                // When the GetCustomizedPageStatus operation is invoked by inValid File Url, 
                // if the server return exception, then the requirement can be captured. 
                // Verify requirement MS-WEBSS_R261
                Site.CaptureRequirement(
                261,
                @"[In GetCustomizedPageStatus] The protocol server portion of the HTTP POST address MUST be the address of a site.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R285
                // When the GetCustomizedPageStatus operation is invoked by inValid File Url, 
                // if the server return exception, then the requirement can be captured. 
                Site.CaptureRequirement(
                285,
                @"[In GetCustomizedPageStatusResponse] Because the fileUrl relative to the protocol server URL does not specify a valid page, the protocol server returns the SOAP fault specified previously.");

                Site.Assert.IsFalse(!Common.IsRequirementEnabled(1067, this.Site), "This operation GetCustomizedPageStatus failed.");
                if (Common.IsRequirementEnabled(1067, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1067
                    // When the GetCustomizedPageStatus operation is invoked by inValid File Url, 
                    // if the server return exception, then the requirement can be captured.
                    Site.CaptureRequirement(
                    1067,
                    @"[In GetCustomizedPageStatusResponse] If implementation does  encounter the following error condition while running this operation[GetCustomizedPageStatus], a SOAP fault with the error code 0x80070002 is returned as follows:
 Occurs when relative to the protocol server address, the value of fileUrl does not specify a valid page.[ The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
                }

                // If the product is Windows® SharePoint® Services 3.0 and above, it will be captured.
                if (Common.IsRequirementEnabled(1030, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R1030
                    // When the GetCustomizedPageStatus operation is invoked by inValid File Url, 
                    // if the server return exception, then the requirement can be captured. 
                    Site.CaptureRequirement(
                        1030,
                        @"[In Appendix B: Product Behavior] Implementation does support this[GetCustomizedPageStatus] operation.(<17> Windows SharePoint Services 3.0 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify WebUrlFromPageUrl operation  to get the URL of the parent site of the specified URL.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S03_TC03_WebUrlFromPageUrl()
        {
            // For capture requirements in adapter.
            bool isVerifiedR633 = false;
            string fullURL = Adapter.WebUrlFromPageUrl(Common.GetConfigurationPropertyValue("WebUrlFromPageUrl_PageUrl", this.Site));
            if (fullURL.Equals(Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site), StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedR633 = true;
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R633
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR633,
                633,
                @"[In WebUrlFromPageUrlResponse] WebUrlFromPageUrlResult: MUST be the full URL of the site.");
        }

        /// <summary>
        /// This test case aims to verify the GetCustomizedPageStatus operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S03_TC04_GetCustomizedPageStatus_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);
            try
            {
                Adapter.GetCustomizedPageStatus(Common.GetConfigurationPropertyValue("GetCustomizedPageStatus_ValidFileUrl", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the GetCustomizedPageStatus operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1077
                // COMMENT: When the GetCustomizedPageStatus operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be 
                // captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1077,
                    @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetCustomizedPageStatus], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case is used to test WebUrlFromPageUrl operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S03_TC05_WebUrlFromPageUrl_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.WebUrlFromPageUrl(Common.GetConfigurationPropertyValue("WebUrlFromPageUrl_PageUrl", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the WebUrlFromPageUrl operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1089
                // COMMENT: When the WebUrlFromPageUrl operation is invoked by unauthenticated user, if 
                // the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1089,
                    @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[WebUrlFromPageUrl], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }
    }
}