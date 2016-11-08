namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using System;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 1 Test cases. Test the CreateSite, DeleteSite and GetLanguages operations.
    /// </summary>
    [TestClass]
    public class S01_CreateAndDeleteSite : TestClassBase
    {
        #region Variables

        /// <summary>
        /// An event handler to add when a test case is start.
        /// </summary>
        private static EventHandler<TestStartFinishEventArgs> specialTestStartedEvent = new EventHandler<TestStartFinishEventArgs>(TestStarted_InvokeSutMethod);

        /// <summary>
        /// An instance of IMS_ADMINSSUTControlAdapter.
        /// </summary>
        private static IMS_ADMINSSUTControlAdapter sutAdapter;

        /// <summary>
        /// A value that identify if the SUT control adapter method "SetUserProfileService" has been invoked successfully. 
        /// </summary>
        private static bool invokeSetUserProfileServiceSuccedd = false;

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
            // Initialize the TestSuiteBase
            TestClassBase.Initialize(testContext);

            // Add the event handler for "TestStarted" event.
            TestClassBase.BaseTestSite.TestStarted += specialTestStartedEvent;
        }

        /// <summary>
        /// Dispose the "TestStarted" event, invoke the SUT control adapter method "SetUserProfileService" with "true" input parameter.
        /// </summary>
        /// <param name="sender">The object of sender</param>
        /// <param name="e">The object represents the event arguments.</param>
        public static void TestStarted_InvokeSutMethod(object sender, TestStartFinishEventArgs e)
        {
            // Remove the event handler for "TestStarted" event. So that the static method is invoked only one time.
            TestClassBase.BaseTestSite.TestStarted -= specialTestStartedEvent;

            TestClassBase.BaseTestSite.GetAdapter<IMS_ADMINSAdapter>();

            S01_CreateAndDeleteSite.sutAdapter = TestClassBase.BaseTestSite.GetAdapter<IMS_ADMINSSUTControlAdapter>();
            TestClassBase.BaseTestSite.Assume.IsNotNull(S01_CreateAndDeleteSite.sutAdapter, "The static object 'sutAdapter' should not be null!");

            bool invokeSutSuccedd = S01_CreateAndDeleteSite.sutAdapter.SetUserProfileService(true);
            TestClassBase.BaseTestSite.Assume.IsTrue(invokeSutSuccedd, "In method 'TestStarted_InvokeSutMethod', the SUT Control Adapter method 'SetUserProfileService' did not invoke successfully! ");

            S01_CreateAndDeleteSite.invokeSetUserProfileServiceSuccedd = invokeSutSuccedd;
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();

            if (S01_CreateAndDeleteSite.invokeSetUserProfileServiceSuccedd == true)
            {
                bool invokeSutSuccedd = S01_CreateAndDeleteSite.sutAdapter.SetUserProfileService(false);
                if (!invokeSutSuccedd)
                {
                    TestClassBase.BaseTestSite.Log.Add(LogEntryKind.Comment, "In method 'ClassCleanup', the SUT Control Adapter method 'SetUserProfileService' did not invoke successfully! ");
                }
            }
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This test case is used to test create and delete the specified site collection with port number.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC01_CreateSiteDeleteSiteSuccessfully_FormatWithPortNumber()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with port number 80.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
           
            // The input parameter URL is an absolute URL of the site collection to be created and the CreateSite method is successful, MS-ADMINS_R13 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                13,
                @"[In CreateSiteSoapIn][The request message is governed by the following rules:]The absolute URL of the site collection to be created MUST be included in the request message.");

            // The input parameter URL is formatted as http://ServerName:PortNumber/sites/SiteCollectionName and the CreateSite operation succeed, MS-ADMINS_R3043 can be captured.
            Site.CaptureRequirementIfIsTrue(
                 Uri.IsWellFormedUriString(result, UriKind.Absolute),
                 3043,
                 @"[In CreateSite]If the Url contains the port number, the Url format is http://ServerName:PortNumber/sites/SiteCollectionName.");

            // The CreateSite operation succeed and the CreateSiteResult is not empty, MS-ADMINS_R1077001 can be captured.
            Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(result),
                1077001,
                @"[In CreateSiteResponse]It[CreateSiteResult] MUST be returned if the CreateSite operation succeeds.");

            // The input parameter URL is formatted as http://ServerName:PortNumber/sites/SiteCollectionName and the CreateSite operation succeed, MS-ADMINS_R2042001 can be captured.
            Site.CaptureRequirement(
                 2042001,
                 @"[In CreateSite]PortNumber in the [Url's] first format [http://ServerName:PortNumber/sites/SiteCollectionName] MUST be the port number used by the web application  on the protocol server.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
            Site.Assert.Pass("The delete site method should succeed.");
           
            // The input parameter URL is an absolute URL of the site collection to be deleted and the DeleteSite method is successful, MS-ADMINS_R83 can be verified.
            Site.CaptureRequirement(
                83,
                @"[In DeleteSite]The [DeleteSiteSoapIn] request message MUST contain the absolute URL of the site collection to be deleted.");

            // The input parameter URL is an absolute URL of the site collection to be deleted and the DeleteSite method is successful, MS-ADMINS_R120 can be verified.
            Site.CaptureRequirement(
                120,
                @"[In DeleteSiteSoapIn][The request message is governed by the following rules:]The absolute URL of the site collection to be deleted MUST be included in the request message.");
        }

        /// <summary>
        /// This test case is used to create the specified site collection with admin port number.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC02_CreateSiteDeleteSiteSuccessfully_FormatWithAdminPortNumber()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with admin port number.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithAdminPort() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
           
            // The input parameter URL is an absolute URL of the site collection to be created and the CreateSite method is successful, MS-ADMINS_R13 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                13,
                @"[In CreateSiteSoapIn][The request message is governed by the following rules:]The absolute URL of the site collection to be created MUST be included in the request message.");

            // The input parameter OwnerLogin is given a valid value and the CreateSite method is successful, MS-ADMINS_R22 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                22,
                @"[In CreateSiteSoapIn]The valid OwnerLogin MUST be included in the request message.");

            // The input parameter URL is formatted as http://ServerName:PortNumber/sites/SiteCollectionName and the CreateSite operation succeed, MS-ADMINS_R3043 can be captured.
            Site.CaptureRequirementIfIsTrue(
                 Uri.IsWellFormedUriString(result, UriKind.Absolute),
                 3043,
                 @"[In CreateSite]If the Url contains the port number, the Url format is http://ServerName:PortNumber/sites/SiteCollectionName.");

            // The input parameter URL is formatted as http://ServerName:PortNumber/sites/SiteCollectionName and the CreateSite operation succeed, MS-ADMINS_R2042002 can be captured.
            Site.CaptureRequirement(
                 2042002,
                 @"[In CreateSite]PortNumber in the [Url's] first format [http://ServerName:PortNumber/sites/SiteCollectionName] MUST be the port number used by or the Administration Web Service on the protocol server.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection without port number.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC03_CreateSiteDeleteSiteSuccessfully_FormatWithoutPortNumber()
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

            // The input parameter URL is an absolute URL of the site collection to be created and the CreateSite method is successful, MS-ADMINS_R13 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                13,
                @"[In CreateSiteSoapIn][The request message is governed by the following rules:]The absolute URL of the site collection to be created MUST be included in the request message.");

            // The input parameter URL is formatted as http://ServerName/sites/SiteCollectionName and the CreateSite operation succeed, MS-ADMINS_R3044 can be captured.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                3044,
                @"[In CreateSite]If the Url does not contain the port number, the Url format is http://ServerName/sites/SiteCollectionName.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with Title element absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC04_CreateSiteSuccessfully_TitleAbsent()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with Title element absent.
            int lcid = lcids.Languages[0];
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + TestSuiteBase.GenerateUniqueSiteTitle();
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, null, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with title absent and the new created site has a default title of "Team Site", then MS-ADMINS_R3017 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "Team Site",
                sutAdapter.GetSiteProperty(result, "Title"),
                3017,
                @"[In CreateSite]If [the Title] nothing is specified, the [created]site will have a default title of ""Team Site"".");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with Description element absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC05_CreateSiteSuccessfully_DescriptionAbsent()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with Description absent.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, null, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with description absent and the new created site does not have a description, then MS-ADMINS_R2046 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                sutAdapter.GetSiteProperty(result, "Description"),
                2046,
                @"[In CreateSite]If [the Description] nothing is specified, the [created] site will not have a description.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with PortalUrl element absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC06_CreateSiteSuccessfully_PortalUrlAbsent()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with PortalUrl absent.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, null, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string defaultPortalUrl = sutAdapter.GetSiteProperty(result, "PortalUrl");

            // If create site successfully with no portalUrl and the returned portalUrl property value equals to empty, so no portal URL set in the database, then MS-ADMINS_R2057 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                defaultPortalUrl,
                2057,
                @"[In CreateSite]If [the PortalUrl] nothing is specified, no portal URL will be set in the database.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with PortalUrl element empty.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC07_CreateSiteSuccessfully_PortalUrlEmpty()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with PortalUrl element empty.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = string.Empty;
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string defaultPortalUrl = sutAdapter.GetSiteProperty(result, "PortalUrl");

            // If create site successfully with empty portalUrl and the returned portalUrl property value equals to empty, so no portal URL set in the database, then MS-ADMINS_R2059 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                defaultPortalUrl,
                2059,
                @"[In CreateSite]If the URL[PortalUrl] is absent or empty, no portal URL will be set in the database.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with PortalName element absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC08_CreateSiteSuccessfully_PortalNameAbsent()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with PortalName absent.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, null);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with portalName absent, and the returned property value of portalName equals to null, so no portal name set in the database, then MS-ADMINS_R2061 can be captured. 
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                sutAdapter.GetSiteProperty(result, "PortalName"),
                2061,
                @"[In CreateSite]If[the PortalName] nothing is specified, no portal name will be set in the database.");
            
            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to test create the specified site collection with owner name exceeding max length.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC09_CreateSiteSuccessfully_OwnerNameExceedMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateRandomString(256);
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite to create a site collection with owner name exceeding max length.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string ownerNameReturned = sutAdapter.GetSiteProperty(result, "OwnerName");

            // If the site collection is created successfully, the OwnerName returned contains 255 characters and the exceeded characters are truncated, then MS-ADMINS_R3024 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                ownerName.Substring(0, 255),
                ownerNameReturned,
                3024,
                @"[In CreateSite]If the length of the OwnerName exceeds 255 characters, the CreateSite operation will succeed  without exception, the exceeds characters are truncated.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with owner email exceeding max length.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC10_CreateSiteSuccessfully_OwnerEmailExceedMaxLength()
        {
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
            string ownerEmail = TestSuiteBase.GenerateEmail(256);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with owner email exceeding max length.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string ownerEmailReturned = sutAdapter.GetSiteProperty(result, "OwnerEmail");

            // If the site collection is created successfully. The OwnerEmail returned contains 255 characters and the exceeded characters are truncated, then MS-ADMINS_R3026 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                ownerEmail.Substring(0, 255),
                ownerEmailReturned,
                3026,
                @"[In CreateSite]If the length of the OwnerEmail exceeds 255 characters, the CreateSite operation will succeed without exception, the exceeded characters are truncated.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with portalUrl exceeding max length.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC11_CreateSiteSuccessfully_PortalUrlExceedMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(261);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with portalUrl exceeding max length.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string portalUrlReturned = sutAdapter.GetSiteProperty(result, "PortalUrl");

            // If the site collection is created successfully. The PortalUrl returned contains 260 characters and the exceeds characters are truncated, then MS-ADMINS_R3028 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                portalUrl.Substring(0, 260),
                portalUrlReturned,
                3028,
                @"[In CreateSite]If the length of the PortalUrl exceeds 260 characters, the CreateSite operation will succeed  without exception, the exceeds characters are truncated.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with portalName exceeding max length.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC12_CreateSiteSuccessfully_PortalNameExceedMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateRandomString(256);

            // Call CreateSite method to create a site collection with portalName exceeding max length.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");
            
            string portalNameReturned = sutAdapter.GetSiteProperty(result, "PortalName");

            // If the site collection is created successfully. The portalName returned contains 255 characters and the exceeded characters are truncated, then MS-ADMINS_R3030 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                portalName.Substring(0, 255),
                portalNameReturned,
                3030,
                @"[In CreateSite]If the length of the PortalName exceeds 255 characters, the CreateSite operation will succeed  without exception, the exceeds characters are truncated.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of title exceeds maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC13_CreateSiteSuccessfully_TitleExceedMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateRandomString(256);
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + TestSuiteBase.GenerateUniqueSiteTitle();
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with the length of title exceeds maximum characters 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string titleReturned = sutAdapter.GetSiteProperty(result, "Title");

            // If create site successfully, and the exceeded part of title length were truncated, the MS-ADMINS_R3034 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                title.Substring(0, 255),
                titleReturned,
                3034,
                @"[In CreateSite]If the length of the Title exceeds 255 characters, the CreateSite operation will succeed without exception, the exceeded characters are truncated.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of title less than maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC14_CreateSiteSuccessfully_TitleLessThanMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateRandomString(254);
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + TestSuiteBase.GenerateUniqueSiteTitle();
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with the length of title less than maximum characters 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with the length of title less than maximum characters 255, then MS-ADMINS_R3015 can be verified.
            Site.CaptureRequirement(
            3015,
                @"[In CreateSite]If the length of the Title is 254 characters, the CreateSite operation will succeed.");
            
            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of title equals to maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC15_CreateSiteSuccessfully_TitleEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateRandomString(255);
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + TestSuiteBase.GenerateUniqueSiteTitle();
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with the length of title equals to maximum characters 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with the length of title equals to maximum characters 255, then MS-ADMINS_R3016 can be verified.
            Site.CaptureRequirement(
                3016,
                @"[In CreateSite]If the length of the Title is 255 characters, the CreateSite operation will succeed.");
            
            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of description less than maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC16_CreateSiteSuccessfully_DescriptionLessThanMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(254);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with description length less than 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of description equals to maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC17_CreateSiteSuccessfully_DescriptionEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with description length equals to 255.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(255);
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
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of description less than maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC18_CreateSiteSuccessfully_OwnerNameLessThanMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateRandomString(254);
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with ownerName length less than 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with ownerName length less than 255, then MS-ADMINS_R3023 can be verified.
            Site.CaptureRequirement(
                3023,
                @"[In CreateSite]If the length of the OwnerName is 254 characters, the CreateSite operation will succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of ownerName equals to maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC19_CreateSiteSuccessfully_OwnerNameEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = Common.GetConfigurationPropertyValue("UrlWithOutPort", this.Site) + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateRandomString(255);
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with ownerName length equals to 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with ownerName length equals to 255, then MS-ADMINS_R3037 can be verified.
            Site.CaptureRequirement(
                3037,
                @"[In CreateSite]If the length of the OwnerName is 255 characters, the CreateSite operation will succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of ownerEmail less than maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC20_CreateSiteSuccessfully_OwnerEmailLessThanMaxLength()
        {
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
            string ownerEmail = TestSuiteBase.GenerateEmail(254);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with ownerEmail length less than 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with ownerEmail length less than 255, and the input ownerEmail equals to the returned ownerEmail property value, then MS-ADMINS_R3025 can be verified.
            Site.CaptureRequirement(
                3025,
                @"[In CreateSite]If the length of the OwnerEmail is 254 characters, the CreateSite operation will succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of ownerEmail equals to maximum characters 255.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC21_CreateSiteSuccessfully_OwnerEmailEqualsToMaxLength()
        {
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
            string ownerEmail = TestSuiteBase.GenerateEmail(255);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite method to create a site collection with ownerEmail length equals to 255.
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with ownerEmail length equals to 255, then MS-ADMINS_R3013 can be verified.
            Site.CaptureRequirement(
                3013,
                @"[In CreateSite]If the length of the OwnerEmail is 255 characters, the CreateSite operation will succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of Url equals to maximum characters 128 not including http://ServerName.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC22_CreateSiteSuccessfully_UrlEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite to create a site collection with the length of Url equals to maximum characters 128 not including http://ServerName.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlWithoutPort(128);
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string urlGet = sutAdapter.GetSiteProperty(result, "Url");

            // If create site successfully, and the input Url equals to the Url property value returned, then MS-ADMINS_R28 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                url.ToLower(CultureInfo.CurrentCulture),
                urlGet.ToLower(CultureInfo.CurrentCulture),
                28,
                @"[In CreateSite]Its[Url's] maximum length, not including http://ServerName or http://ServerName:PortNumber, is 128 characters.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of PortalUrl less than maximum characters 260. 
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC23_CreateSiteSuccessfully_PortalUrlLessThanMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with the length of PortalUrl less than maximum characters 260.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(259);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // If create site successfully with PortalUrl length equals to 259, than MS-ADMINS_R3027 can be verified.
            Site.CaptureRequirement(
                3027,
                @"[In CreateSite]If the length of the PortalUrl is 259 characters, the CreateSite operation will succeed.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of PortalUrl equals to maximum characters 260. 
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC24_CreateSiteSuccessfully_PortalUrlEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with the length of PortalUrl equals to maximum characters 260.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(260);
            string portalName = TestSuiteBase.GenerateUniquePortalName();
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string portalUrlGet = sutAdapter.GetSiteProperty(result, "PortalUrl");

            // If create site successfully with PortalUrl length equals to 260, then MS-ADMINS_R3038 can be verified.
            Site.CaptureRequirement(
                3038,
                @"[In CreateSite]If the length of the PortalUrl is 260 characters, the CreateSite operation will succeed.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of portalName less than maximum characters 255. 
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC25_CreateSiteSuccessfully_PortalNameLessThanMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with the length of portalName less than maximum characters 255.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateRandomString(254);
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string portalNameGet = sutAdapter.GetSiteProperty(result, "PortalName");

            // If create site successfully with portalName length equals to 254, than MS-ADMINS_R3029 can be verified.
            Site.CaptureRequirement(
                3029,
                @"[In CreateSite]If the length of the PortalName is 254 characters, the CreateSite operation will succeed.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the length of portalName equals to maximum characters 255. 
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC26_CreateSiteSuccessfully_PortalNameEqualsToMaxLength()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with the length of portalName equals to maximum characters 255.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateRandomString(255);
            
            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string portalNameGet = sutAdapter.GetSiteProperty(result, "PortalName");

            // If create site successfully with portalName length equals to 255 then MS-ADMINS_R3039 can be verified.
            Site.CaptureRequirement(
                3039,
                @"[In CreateSite]If the length of the PortalName is 255 characters, the CreateSite operation will succeed.");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection with the LCID element absent.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC27_CreateSiteSuccessfully_LcidAbsent()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3041, this.Site), "This case runs only when the requirement 3041 is enabled.");
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string webTemplate = Common.GetConfigurationPropertyValue("CustomizedTemplate", this.Site);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            // Call CreateSite to create a site collection with the LCID element absent.                
            string result = this.adminsAdapter.CreateSite(url, title, description, null, webTemplate, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);

            // If CreateSite succeed and the returned value is a well formatted URL. That is to say, the server assigned a default value when the LCID element is missing. MS-ADMINS_R3041 can be verified.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(result, UriKind.Absolute),
                3041,
                @"[In CreateSite]Implementation does assign a default LCID based on the default server install language when it is missing. [In Appendix B: Product Behavior](<1> Section 3.1.4.1.2.1: Windows SharePoint Services 3.0 and above follow this behavior)");

            // Call DeleteSite to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection without optional parameters.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC28_CreateSiteSuccessfully_WithRequiredParameters()
        {
            // Call CreateSite method to create a site without optional parameters.          
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + TestSuiteBase.GenerateRandomString(10);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            
            string result = this.adminsAdapter.CreateSite(url, null, null, null, null, ownerLogin, null, null, null, null);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
        }

        /// <summary>
        /// This test case is used to create the specified site collection without template.
        /// </summary>
        [TestCategory("MSADMINS"), TestMethod()]
        public void MSADMINS_S01_TC29_CreateSiteSuccessfully_WithoutTemplate()
        {
            // Call GetLanguages method to obtain LCID values used in the protocol server deployment. 
            GetLanguagesResponseGetLanguagesResult lcids = this.adminsAdapter.GetLanguages();
            Site.Assert.IsNotNull(lcids, "Get languages should succeed and a list of LCIDs should return. If no LCID returns the get languages method is failed.");

            // Call CreateSite method to create a site collection with port number 80.
            int lcid = lcids.Languages[0];
            string title = TestSuiteBase.GenerateUniqueSiteTitle();
            string url = TestSuiteBase.GenerateUrlPrefixWithPortNumber() + title;
            string description = TestSuiteBase.GenerateRandomString(20);
            string ownerLogin = Common.GetConfigurationPropertyValue("OwnerLogin", this.Site);
            string ownerName = TestSuiteBase.GenerateUniqueOwnerName();
            string ownerEmail = TestSuiteBase.GenerateEmail(20);
            string portalUrl = TestSuiteBase.GeneratePortalUrl(20);
            string portalName = TestSuiteBase.GenerateUniquePortalName();

            string result = this.adminsAdapter.CreateSite(url, title, description, lcid, null, ownerLogin, ownerName, ownerEmail, portalUrl, portalName);
            Site.Assert.IsTrue(Uri.IsWellFormedUriString(result, UriKind.Absolute), "Create site should succeed.");

            string siteProperty = sutAdapter.GetSiteProperty(result, "Configuration");

            //  if no template is specified, the "Configuration" is a invalid value "-1". MS-ADMINS_R2048 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "-1",
                siteProperty,
                2048,
                @"[In CreateSite]If no template[WebTemplate] is specified, then no template will be applied to the site at creation time.");


            // Call DeleteSite method to delete the site collection created in above steps.
            this.adminsAdapter.DeleteSite(result);
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
            this.adminsAdapter = this.Site.GetAdapter<IMS_ADMINSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            if (S01_CreateAndDeleteSite.invokeSetUserProfileServiceSuccedd == false)
            {
                this.Site.Assume.Inconclusive("The test environment is not ready to run the test case, because the SUT control adapter method SetUserProfileService could be not invoked successfully.");
            }

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