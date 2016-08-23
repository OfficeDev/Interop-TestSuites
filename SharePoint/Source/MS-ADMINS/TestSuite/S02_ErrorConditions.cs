namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using System;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 2 Test cases. Test the CreateSite, DeleteSite and GetLanguages operations error conditions.
    /// </summary>
    [TestClass]
    public class S02_ErrorConditions : TestClassBase
    {
        #region Variables

        /// <summary>
        /// An instance of IMSADMINSAdapter.
        /// </summary>
        private IMS_ADMINSAdapter adminsAdapter;

        #endregion

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        /// <param name="testContext">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            // Setup test site.
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // Cleanup test site, must be called to ensure closing of logs.
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This test case is used to create the specified site collection with URL exceeding the max length.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC01_CreateSiteFailed_UrlExceedMaxLength()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string urlExceedMaxLength = TestSuiteBase.GenerateUrlWithoutPort(129);
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method to create a site collection with Url exceeding the max length.
                this.adminsAdapter.CreateSite(urlExceedMaxLength, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when set the length of Url exceeding 128 characters, MS-ADMINS_R1028 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R1028
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                1028,
                @"[In CreateSiteSoapIn]If Url's length not including ""http://ServerName""  exceeds 128 characters, the server  MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with URL server name invalid.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC02_CreateSiteFailed_UrlServerNameInvalid()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string urlServerNameInvalid = Common.GetConfigurationPropertyValue("TransportType", this.Site) + "://" + TestSuiteBase.GenerateRandomString(5) + "/sites/" + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {   // Call CreateSite method to create a site collection with invalid Url server name.
                this.adminsAdapter.CreateSite(urlServerNameInvalid, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when set the server name invalid, MS-ADMINS_R1024 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R1024
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                1024,
                @"[In CreateSiteSoapIn]If ServerName in the URL is invalid, the server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with invalid port number.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC03_CreateSiteFailed_UrlPortNumberInvalid()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string urlPortNumberInvalid = Common.GetConfigurationPropertyValue("TransportType", this.Site) + "://" + Common.GetConfigurationPropertyValue("SutComputerName", this.Site) + ":" + Common.GetConfigurationPropertyValue("InvalidPortNumber", this.Site) + "/sites/" + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method to create a site collection with invalid port number.
                this.adminsAdapter.CreateSite(urlPortNumberInvalid, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when set the port number invalid, MS-ADMINS_R1025 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R1025
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                1025,
                @"[In CreateSiteSoapIn]If the PortNumber in the URL given an invalid value, the server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with invalid URL format.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC04_CreateSiteFailed_UrlInvalidFormat()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string urlInvalidFormat = Common.GetConfigurationPropertyValue("TransportType", this.Site) + "://" + "sites/" + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method to create a site collection with invalid Url format.
                this.adminsAdapter.CreateSite(urlInvalidFormat, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when set the url format invalid, MS-ADMINS_R1026 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R1026
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                1026,
                @"[In CreateSiteSoapIn]If the URL does not comply with either of the two formats: http://ServerName:PortNumber/sites/SiteCollectionName or http://ServerName/sites/SiteCollectionName, the server  MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with URL already existed.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC05_CreateSiteFailed_UrlExisted()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];

            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string titleSnd = TestSuiteBase.GenerateUniqueSiteTitle();
            string descriptionSnd = TestSuiteBase.GenerateRandomString(30);
            string webTemplateSnd = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLoginSnd = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerNameSnd = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmailSnd = TestSuiteBase.GenerateEmail(20);
            string portalUrlSnd = TestSuiteBase.GeneratePortalUrl(20);
            string portalNameSnd = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method again to create a site collection with existed Url.
                this.adminsAdapter.CreateSite(url, titleSnd, descriptionSnd, lcid, webTemplateSnd, ownerLoginSnd, ownerNameSnd, ownerEmailSnd, portalUrlSnd, portalNameSnd);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when give an existed Url, MS-ADMINS_R17 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R17
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                17,
                @"[In CreateSiteSoapIn][The request message is governed by the following rules:]If the URL already exists, the server MUST return a SOAP fault.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with URL absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC06_CreateSiteFailed_UrlAbsent()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method to create a site collection with Url absent.
                this.adminsAdapter.CreateSite(null, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when Url absent, MS-ADMINS_R14 and MS-ADMINS_R2041 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R14
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                14,
                @"[In CreateSiteSoapIn][The request message is governed by the following rules:]If the URL is missing, the server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                2041,
                @"[In CreateSite]If it[Url] is missing or absent, the server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with URL empty.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC07_CreateSiteFailed_UrlEmpty()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = string.Empty;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite method to create a site collection with Url empty.
                this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            Site.Log.Add(LogEntryKind.Debug, "If the Soap fault returned when Url empty, MS-ADMINS_R1027 can be verified.");

            // Verify MS-ADMINS requirement: MS-ADMINS_R1027
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                1027,
                @"[In CreateSiteSoapIn]If the URL  is empty, the server  MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection without installed LCID.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC08_CreateSiteFailed_LcidNotInstalled()
        {
            string strErrorCode = string.Empty;
            int notInstalledLcid = int.Parse(Common.GetConfigurationPropertyValue("NotInstalledLCID", this.Site));
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            try
            {
                // Call CreateSite method to create a site collection with a not installed LCID.                
                this.adminsAdapter.CreateSite(url, title, description, notInstalledLcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException exp)
            {
                strErrorCode = Common.ExtractErrorCodeFromSoapFault(exp);
            }

            // If the returned error code equals to 0x8102005e, then MS-ADMINS_R18 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x8102005e",
                strErrorCode,
                18002,
                @"[In CreateSiteSoapIn] If the LCID is invalid [or not installed], then the server MUST return a SOAP fault with error code 0x8102005e.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with an invalid LCID.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC09_CreateSiteFailed_LcidInvalid()
        {
            string strErrorCode = string.Empty;
            int invalidLcid = TestSuiteBase.GenerateRandomNumber(0, 99);
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            try
            {
                // Call CreateSite to create a site collection with an invalid LCID.                
                this.adminsAdapter.CreateSite(url, title, description, invalidLcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException exp)
            {
                strErrorCode = Common.ExtractErrorCodeFromSoapFault(exp);
                Site.Log.Add(LogEntryKind.Debug, "Soap exception returned, it means the create site operation failed with invalid LCID inputting, the message is: {0}", exp.Message);
            }

            // If the returned error code equals to 0x8102005e, then MS-ADMINS_R18 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x8102005e",
                strErrorCode,
                18,
                @"[In CreateSiteSoapIn] If the LCID is invalid [or not installed], then the server MUST return a SOAP fault with error code 0x8102005e.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with invalid WebTemplate.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC10_CreateSiteFailed_WebTemplateInvalid()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplateInvalid = TestSuiteBase.GenerateRandomString(5);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            try
            {
                // Call CreateSite method to create a site collection with invalid WebTemplate.                
                this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplateInvalid, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R19 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                19,
                @"[In CreateSiteSoapIn]If WebTemplate is not empty, and if it is not available in the list of templates and it is not a custom template, then the server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with owner login name not existed.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC11_CreateSiteFailed_OwnerLoginAccountNotExisted()
        {
            string strErrorCode = string.Empty;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLoginNotExisted = TestSuiteBase.GenerateRandomString(10);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            try
            {
                // Call the CreateSite method to create a site collection with owner login name not existed.
                this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLoginNotExisted, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException exp)
            {
                strErrorCode = Common.ExtractErrorCodeFromSoapFault(exp);
            }

            // If the returned error code equals to 0x80131600, then MS-ADMINS_R1022 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x80131600",
                strErrorCode,
                1022,
                @"[In CreateSiteSoapIn]If OwnerLogin is not an existing domain user account, the server MUST return a SOAP fault with error code 0x80131600.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection without owner login.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC12_CreateSiteFailed_OwnerLoginAbsent()
        {
            bool isSoapFaultReturn = false;

            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            try
            {
                // Call CreateSite method to create a site collection without owner login.
                this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, null, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R2050 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                2050,
                @"[In CreateSite]If it[OwnerLogin] is missing, the server MUST return a SOAP fault.");

            isSoapFaultReturn = false;
            try
            {
                // Call CreateSite method to create a site collection with empty owner login.
                this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, string.Empty, ownerName, ownerEmail, portalUrl, portalName);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R2050 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                2050001,
                @"[In CreateSite]If it[OwnerLogin] is empty, the server MUST return a SOAP fault.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with empty ownerLogin.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC13_CreateSiteFailed_OwnerLoginEmpty()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLoginEmpty = string.Empty;
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            
            try
            {
                // Call CreateSite to create a site collection with empty ownerLogin.
                string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLoginEmpty, ownerName, ownerEmail, portalUrl, portalName);
                Site.Assert.IsFalse(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site operation should fail.");
            }
            catch (SoapException)
            {
                Site.Log.Add(LogEntryKind.Debug, "Create site operation should fail with empty OwnerLogin.");
            }
        }

        /// <summary>
        /// This test case is used to delete the site without URL specified.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC14_DeleteSiteFailed_UrlMissing()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite to create a site collection without port number.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            bool isSoapFaultReturn = false;
            try
            {
                // Call DeleteSite method without URL specified.
                this.adminsAdapter.DeleteSite(null);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R121 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                121,
                @"[In DeleteSiteSoapIn][The request message is governed by the following rules:] If the URL is missing, the server MUST return a SOAP fault.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to delete the site collection with invalid URL (using server name without port number as an example).
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC15_DeleteSiteFailed_UrlNameInvalid()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection without port number.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            bool isSoapFaultReturn = false;
            try
            {
                string invalidUrl = Common.GetConfigurationPropertyValue("TransportType", this.Site) + TestSuiteBase.GenerateRandomString(5) + "/sites/" + title;

                // Call DeleteSite method with invalid URL.
                this.adminsAdapter.DeleteSite(invalidUrl);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R122 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                122,
                @"[In DeleteSiteSoapIn][The request message is governed by the following rules:] If the URL is not valid, the server MUST return a SOAP fault.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to delete the site collection with a nonexistent URL.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S02_TC16_DeleteSiteFailed_UrlNotExist()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection without port number.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
         
            bool isSoapFaultReturn = false;
            try
            {
                // Call DeleteSite method with a nonexistent URL.
                this.adminsAdapter.DeleteSite(result);
            }
            catch (SoapException)
            {
                isSoapFaultReturn = true;
            }

            // If a SOAP fault is returned, then MS-ADMINS_R123 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultReturn,
                123,
                @"[In DeleteSiteSoapIn][The request message is governed by the following rules:] If the URL does not exist, the server MUST return a SOAP fault.");
        }

        #endregion

        #region Test Case Initialization and Cleanup

        /// <summary>
        /// Overrides TestClassBase's TestInitialize().
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            // Initialization of adapter.
            this.adminsAdapter = Site.GetAdapter<IMS_ADMINSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            // Initialize the TestSuiteBase
            TestSuiteBase.Initialize(this.Site);
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup().
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            // Resetting of adapter.
            this.adminsAdapter.Reset();
        }

        #endregion
    }
}