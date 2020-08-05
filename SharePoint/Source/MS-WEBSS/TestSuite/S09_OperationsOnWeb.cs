namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{ 
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with sub-webs, webs and web collection.
    /// </summary>
    [TestClass]
    public class S09_OperationsOnWeb : TestSuiteBase
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

        #region GetAllSubWebCollection Test Cases

        /// <summary>
        /// This test case aims to verify the GetWeb operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC01_GetWeb_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetWeb(Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site));
                Site.Assert.Fail("The expected http status code is not returned for the GetWeb operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1080
                // COMMENT: When the GetWeb operation is invoked by unauthenticated user, if the 
                // server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1080,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetWeb], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify GetWeb operation to get the Title, URL, Description, Language, and theme properties of the specified site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC02_GetWeb_Succeed()
        {
            string webUrl = Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site);
            GetWebResponseGetWebResult getWebResult = Adapter.GetWeb(webUrl);

            Site.Assert.IsTrue(
                getWebResult.Web.Title.Equals(Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site), StringComparison.CurrentCultureIgnoreCase),
                "The web title is {0}",
                getWebResult.Web.Title);

            Site.Assert.IsTrue(
                getWebResult.Web.Url.Equals(webUrl, StringComparison.OrdinalIgnoreCase),
                "The web URL is {0}",
                getWebResult.Web.Url);

            Site.Assert.IsTrue(
                getWebResult.Web.Description.Equals(Common.GetConfigurationPropertyValue("TestSiteDescription", this.Site), StringComparison.OrdinalIgnoreCase),
                "The web description is {0}",
                getWebResult.Web.Description);

            Site.Assert.IsTrue(
                getWebResult.Web.Language.Equals(Common.GetConfigurationPropertyValue("TestSiteLanguage", this.Site), StringComparison.OrdinalIgnoreCase),
                "The web Language is {0}",
                getWebResult.Web.Language);

            // Verify MS-WEBSS requirement: MS-WEBSS_R347
            bool isVerifiedGetWeb = false;
            bool titleOfTest = getWebResult.Web.Title.Equals(Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site), StringComparison.OrdinalIgnoreCase);
            bool testSiteDescription = getWebResult.Web.Description.Equals(Common.GetConfigurationPropertyValue("TestSiteDescription", this.Site), StringComparison.OrdinalIgnoreCase);
            bool testSiteLanguage = getWebResult.Web.Language.Equals(Common.GetConfigurationPropertyValue("TestSiteLanguage", this.Site), StringComparison.OrdinalIgnoreCase);
            bool webURL = getWebResult.Web.Url.Equals(webUrl, StringComparison.OrdinalIgnoreCase);
            bool theme = getWebResult.Web.Theme == null || getWebResult.Web.Theme != null;
            isVerifiedGetWeb = titleOfTest && testSiteDescription && testSiteLanguage && webURL && theme;
            Site.CaptureRequirementIfIsTrue(
                isVerifiedGetWeb,
                347,
                @"[In GetWeb] If the operation succeeds, the protocol server MUST return the title, URL, description, language, and theme properties of the specified site.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R950
            Site.CaptureRequirementIfIsTrue(
                isVerifiedGetWeb,
                950,
                @"[In Messages] [Message] GetWebSoapOut [Description] 
The response to the request for the title, URL, description, language, and theme of the specified site (2).");

            // Verify MS-WEBSS requirement: MS-WEBSS_R353
            Site.CaptureRequirementIfIsTrue(
                isVerifiedGetWeb,
                353,
                @"[In GetWebSoapOut] This[GetWebSoapOut] message is the response containing the title, URL, description, language, and theme of the specified site.");
        }

        /// <summary>
        /// This test case aims to verify the GetWeb operation when an invalid URL is passed in the site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC03_GetWeb_InvalidUrl()
        {
            string webUrl = this.GenerateRandomString(10);

            string soapErrorCode = string.Empty;
            try
            {
                Adapter.GetWeb(webUrl);
                Site.Assert.Fail("The expected error code is not returned for the GetWeb operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2-1/2007] section 5.4, SOAP Fault.");

                // Verify MS-WEBSS requirement: MS-WEBSS_R350       
                Site.CaptureRequirementIfIsNotNull(
                    exp,
                    350,
                    @"[In GetWeb] Errors specific to this operation[GetWeb] are defined with GetWebSoapOut message.");

                // Catch the exception, then the following requirements will be captured.
                if ((exp.Detail.LastChild.Name == ConstString.ErrorCode)
                    && !string.IsNullOrEmpty(exp.Detail.LastChild.InnerText))
                {
                    soapErrorCode = exp.Detail.LastChild.InnerText;

                    if ((exp.Detail.FirstChild.Name == ConstString.ErrorString)
                        && !string.IsNullOrEmpty(exp.Detail.FirstChild.InnerText))
                    {
                        // Verify MS-WEBSS requirement: MS-WEBSS_R349
                        // COMMENT: When the GetWeb operation is invoked with invalid URL passed to the site, 
                        // if the response contains ErrorCode element and ErrorString element and the values of 
                        // the two elements are not null or empty, then the requirement can be captured.
                        Site.CaptureRequirement(
                            349,
                            @"[In GetWeb]  If there is any problem in performing the operation, the protocol server MUST return the appropriate error code and error string.");
                    }

                    if (Common.IsRequirementEnabled(743, this.Site))
                    {
                        // COMMENT: When the GetWeb operation is invoked with invalid URL passed to the site, 
                        // if the returned error code is 0x82000001, then the requirement can be captured.
                        // Verify MS-WEBSS requirement: MS-WEBSS_R743
                        bool isR743Satisfied = soapErrorCode == SoapErrorCode.ErrorCode0x82000001;
                        Site.Assert.AreEqual<string>(
                            SoapErrorCode.ErrorCode0x82000001,
                            soapErrorCode,
                            "The SoapErrorCode is {0}",
                            soapErrorCode);

                        // Verify MS-WEBSS requirement: MS-WEBSS_R743
                        Site.CaptureRequirementIfAreEqual<bool>(
                        true,
                        isR743Satisfied,
                        743,
                        @"[In GetWebResponse] Implementation does encounter the following error condition when running this operation, a SOAP fault with the error code 0x82000001 does be returned: Occurs when an invalid URL is passed in the site.
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
                    }
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the GetWeb operation when no webUrl is supplied for.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC04_GetWeb_NoWebUrl()
        {
            string soapErrorCode = string.Empty;
            try
            {
                Adapter.GetWeb(null);
                Site.Assert.Fail("The expected error code is not returned for the GetWeb operation.");
            }
            catch (SoapException exp)
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1042
                // COMMENT: If the server return the server faults as SOAP faults, then the requirement 
                // can be captured.
                Site.CaptureRequirement(
                    1042,
                    @"[In Transport] Protocol server faults MUST be returned SOAP faults as specified either in [SOAP1.1] section 4.4, SOAP Fault, or in [SOAP1.2-1/2007] section 5.4, SOAP Fault.");

                // Catch the exception, then the following requirements will be captured.
                if ((exp.Detail.LastChild.Name == ConstString.ErrorCode)
                    && !string.IsNullOrEmpty(exp.Detail.LastChild.InnerText))
                {
                    soapErrorCode = exp.Detail.LastChild.InnerText;

                    if ((exp.Detail.FirstChild.Name == ConstString.ErrorString)
                        && !string.IsNullOrEmpty(exp.Detail.FirstChild.InnerText))
                    {
                        // Verify MS-WEBSS requirement: MS-WEBSS_R349
                        // COMMENT: When the GetWeb operation is invoked with invalid URL passed to the site, 
                        // if the response contains ErrorCode element and ErrorString element and the values of 
                        // the two elements are not null or empty, then the requirement can be captured.
                        Site.CaptureRequirement(
                            349,
                            @"[In GetWeb]  If there is any problem in performing the operation, the protocol server MUST return the appropriate error code and error string.");
                    }
                }

                if (Common.IsRequirementEnabled(744, this.Site))
                {
                    // Verify MS-WEBSS requirement: MS-WEBSS_R744
                    // COMMENT: When the GetWeb operation is invoked with invalid URL passed to the site, 
                    // if the returned error code is 0x82000001, then the requirement can be captured.
                    bool isR744Satisfied = soapErrorCode == SoapErrorCode.ErrorCode0x82000001;

                    Site.Assert.AreEqual<string>(
                        SoapErrorCode.ErrorCode0x82000001,
                        soapErrorCode,
                        "The SoapErrorCode is {0}",
                        soapErrorCode);

                    // Verify MS-WEBSS requirement: MS-WEBSS_R744
                    Site.CaptureRequirementIfIsTrue(
                    isR744Satisfied,
                    744,
             @"[In GetWebResponse]Implementation does encounter the following error condition when running this operation, a SOAP fault with the error code 0x82000001 does be returned: Occurs when no webUrl element is supplied for the site.
[The 2007 Microsoft® Office system
  Microsoft® Office 2010 suites
  Microsoft® Office SharePoint® Server 2007
  Windows® SharePoint® Services 3.0
  Microsoft® SharePoint® Foundation 2010
Microsoft® SharePoint® Foundation 2013]");
                }
            }
        }

        /// <summary>
        /// This test case aims to verify the GetWebCollection operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC05_GetWebCollection_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetWebCollection();
                Site.Assert.Fail("The expected http status code is not returned for the GetWebCollection operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1081
                // COMMENT: When the GetWebCollection operation is invoked by unauthenticated user, 
                // if the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1081,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetWebCollection], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case verifies the GetWebCollection operation succeeds.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC06_GetWebCollection_Succeed()
        {
            GetWebCollectionResponseGetWebCollectionResult getWebCollectionResult = Adapter.GetWebCollection();

            bool isVerifiedTitleAndUrl = false;
            if (getWebCollectionResult.Webs.Length > 0)
            {
                foreach (WebDefinition w in getWebCollectionResult.Webs)
                {
                    if (w.Url.Equals(Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site), StringComparison.OrdinalIgnoreCase) && w.Title.Equals(Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site), StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifiedTitleAndUrl = true;
                    }

                    if (string.IsNullOrEmpty(w.Url) || string.IsNullOrEmpty(w.Title))
                    {
                        isVerifiedTitleAndUrl = false;
                        Site.Assert.Fail("It does not return expected Title or Url.");
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedTitleAndUrl,
                368,
                @"[In GetWebCollection] If the operation succeeds, it MUST return the Title and URL properties of all immediate child sites of the context site.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R373
            Site.CaptureRequirementIfIsTrue(
               isVerifiedTitleAndUrl,
                373,
                @"[In GetWebCollectionSoapOut] This message[GetWebCollectionSoapOut] is the response containing the Title and URL properties of all immediate child sites of the context site.");
        }

        /// <summary>
        /// This test case aims to verify the GetAllSubWebCollection operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC07_GetAllSubWebCollection_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetAllSubWebCollection();
                Site.Assert.Fail("The expected http status code is not returned for the GetAllSubWebCollection operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1073
                // COMMENT: When the GetAllSubWebCollection operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be 
                // captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1073,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetAllSubWebCollection], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary>
        /// This test case aims to verify GetAllSubWebCollection operation when the operation is succeed.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC08_GetAllSubWebCollection_Succeed()
        {
            GetAllSubWebCollectionResponseGetAllSubWebCollectionResult getAllSubWebResult = Adapter.GetAllSubWebCollection();
            this.Site.Assert.IsNotNull(getAllSubWebResult, "GetAllSubWebCollection operation should succeed.");
        }

        /// <summary>
        /// This test case aims to call GetWeb operation to verify that the WebDefinition complex type that specifies a single site.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC09_GetWebDefinition()
        {
            string webUrl = Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site);
            GetWebResponseGetWebResult getWebResult = Adapter.GetWeb(webUrl);

            bool isVerifiedWebResult = false;
            string titleOfTest = Common.GetConfigurationPropertyValue("TestSiteTitle", this.Site);
            if (getWebResult.Web.Title.Equals(titleOfTest, StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedWebResult = true;
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R857
            Site.CaptureRequirementIfIsTrue(
            isVerifiedWebResult,
            857,
            @"[In WebDefinition] Title: Specifies the title of the site.");

            bool isVerified693 = false;
            if (getWebResult.Web.Url.Equals(webUrl, StringComparison.OrdinalIgnoreCase))
            {
                isVerified693 = true;
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R693
            Site.CaptureRequirementIfIsTrue(
                isVerified693,
                693,
                @"[In WebDefinition] Url: Specifies the absolute URL of the site.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R694
            bool isVerifiedR694 = false;
            string testSiteDescription = Common.GetConfigurationPropertyValue("TestSiteDescription", this.Site);
            if (getWebResult.Web.Description.Equals(testSiteDescription, StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedR694 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR694,
                694,
                @"[In WebDefinition] Description: Specifies the description of the site.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R695
            string testSiteLanguage = Common.GetConfigurationPropertyValue("TestSiteLanguage", this.Site);
            Site.CaptureRequirementIfAreEqual<string>(
                testSiteLanguage,
                getWebResult.Web.Language,
                695,
                @"[In WebDefinition] Language: Specifies the language code identifier (LCID) for the language of the site.");
        }

        /// <summary>
        /// This test case aims to call GetWeb operation to verify that the SOAP body contains a GetWebResponse element.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S09_TC10_GetWebResponseWithValidXml()
        {
            string webUrl = Common.GetConfigurationPropertyValue("SubSiteUrl", this.Site);
            GetWebResponseGetWebResult getWebResult = Adapter.GetWeb(webUrl);

            // Verify MS-WEBSS requirement: MS-WEBSS_R361
            bool isVerifiedWebResult = false;
            if (getWebResult.Web != null)
            {
                isVerifiedWebResult = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedWebResult,
                361,
                @"[In GetWebResponse] GetWebResult: An XML element that contains a Web element.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R362
            Site.CaptureRequirementIfIsTrue(
                isVerifiedWebResult,
                362,
                @"[In GetWebResponse] Web: The structure of this element is defined in WebDefinition (section 2.2.4.2).");
        }
    }
        #endregion
}