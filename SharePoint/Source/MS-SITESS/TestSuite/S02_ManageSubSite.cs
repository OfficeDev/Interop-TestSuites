namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Collections.Generic;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class contains test case definitions related to GetSiteTemplates, CreateWeb and DeleteWeb operations.
    /// </summary>
    [TestClass]
    public class S02_ManageSubSite : TestClassBase
    {
        /// <summary>
        /// An instance of protocol adapter class.
        /// </summary>
        private IMS_SITESSAdapter sitessAdapter;

        /// <summary>
        /// An instance of SUT control adapter class.
        /// </summary>
        private IMS_SITESSSUTControlAdapter sutAdapter;

        /// <summary>
        /// The name of the sub site to be imported.
        /// </summary>
        private string newSubsite = string.Empty;

        #region Test Suite Initialization & Cleanup

        /// <summary>
        /// Test Suite Initialization.
        /// </summary>
        /// <param name="testContext">The test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Test Suite Cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        #region Scenario 2 Manage a subsite

        /// <summary>
        /// This test case is designed to verify GetSiteTemplates, CreateWeb and DeleteWeb operations when managing a subsite successfully.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC01_ManagingSubsiteSuccessfully()
        {
            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + this.newSubsite;
            CreateWebResponseCreateWebResult createResult;
            string expectedUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site)
                + "/"
                + this.newSubsite;
            uint getTemplateResult = 0;
            string webName = this.newSubsite;

            #endregion Variables

            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // If the templateList is not empty, it means the GetSiteTemplates operation is succeed.
            Site.Assert.IsTrue(
                 templateList != null && templateList.Length > 1,
                "GetTemplate operation should return more than one template.");

            #region Capture requirements

            this.VerifyResultOfGetSiteTemplate(getTemplateResult);

            #endregion Capture requirements

            // Invoke the CreateWeb operation with valid parameters, so the return value is expected to contain a URL consistent with the expected URL.
            // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
            createResult = this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateList[1].Name, localeId, true, localeId, true, localeId, true, true, true, true, true, true, true);
            expectedUrl = expectedUrl.TrimEnd('/');
            string actualUrl = createResult.CreateWeb.Url.TrimEnd('/');

            // If returned value contain a URL consistent with the expected URL, it means the CreateWeb operation succeed.
            Site.Assert.IsTrue(
                expectedUrl.Equals(actualUrl, StringComparison.CurrentCultureIgnoreCase),
                "Created web's url should be {0}, but the actual value is {1}.",
                expectedUrl,
                actualUrl);

            #region Capture requirements
            // Get a string contains the name and value of the expected properties of the created web.
            string webPropertyDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), webName);
            //// Get each property value by splitting the string.
            Dictionary<string, string> properties = AdapterHelper.DeserializeWebProperties(webPropertyDefault, Constants.ItemSpliter, Constants.KeySpliter);
            string permissionActual = properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyUserNameInPermissions, this.Site)];
            string currentUser = properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyCurrentUser, this.Site)];
            bool anonymousActual = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyAnonymous, this.Site)]);
            bool presence = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyPresence, this.Site)]);


            // If uniquePermissions is true, R518 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R518, the default user is {0}", currentUser);

            // Verify MS-SITESS requirement: MS-SITESS_R518
            bool isVerifyR518 = false;
            if (permissionActual.Equals(currentUser, StringComparison.CurrentCultureIgnoreCase))
            {
                isVerifyR518 = true;
            }
            else
            {
                string uniqueUserName = permissionActual;
                string currentUserName = currentUser;
                if (permissionActual.Contains("\\"))
                {
                    int index = permissionActual.IndexOf('\\');
                    uniqueUserName = permissionActual.Substring(index + 1, permissionActual.Length - index - 1);
                }

                if (currentUser.Contains("\\"))
                {
                    int index = currentUser.IndexOf('\\');
                    currentUserName = currentUser.Substring(index + 1, currentUser.Length - index - 1);
                }

                if (uniqueUserName.Equals(currentUserName, StringComparison.CurrentCultureIgnoreCase))
                {
                    isVerifyR518 = true;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR518,
                518,
                @"[In CreateWeb] uniquePermissions: If set to true, specifies that the subsite to be created uses its own set of permissions, which default to the current user having full control and no other users having access");

            // If anonymous is true, R520 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R520, the anonymousActual is {0}", anonymousActual);

            // Verify MS-SITESS requirement: MS-SITESS_R520
            Site.CaptureRequirementIfIsTrue(
                anonymousActual,
                520,
                @"[In CreateWeb] anonymous: If set to true, the anonymous authentication is to be allowed for the subsite to be created.");

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R522, the presence is {0}", presence);

            // Verify MS-SITESS requirement: MS-SITESS_R522
            Site.CaptureRequirementIfIsTrue(
                presence,
                522,
                @"[In CreateWeb] presence: If set to true, the online presence information is to be enabled for the subsite to be created.");

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            this.VerifyOperationCreateWeb();
            #endregion Capture requirements

            // If R3781 is not enabled, that means the CreateWeb operation is not supported, so there is no web to be deleted here.
            if (Common.IsRequirementEnabled(3781, this.Site) && Common.IsRequirementEnabled(3791, this.Site))
            {
                // Invoke the DeleteWeb operation.
                this.sitessAdapter.DeleteWeb(webUrl);

                #region Capture requirements
                //// Verify that Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.
                this.VerifyOperationDeleteWeb();
                #endregion Capture requirements
            }
        }

        /// <summary>
        /// This test case is designed to verify GetSiteTemplates, CreateWeb and DeleteWeb operations when managing a subsite without optional parameters.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC02_ManagingSubsiteWithoutOptionalParameters()
        {
            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + this.newSubsite;
            CreateWebResponseCreateWebResult createResult;
            string expectedUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site)
                + "/"
                + this.newSubsite;
            uint getTemplateResult = 0;
            string webName = this.newSubsite;
            uint language = uint.Parse(Common.GetConfigurationPropertyValue(Constants.DefaultLCID, this.Site));
            uint locale = uint.Parse(Common.GetConfigurationPropertyValue(Constants.DefaultLCID, this.Site));

            #endregion Variables

            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // GetTemplate operation should return more than one template.
            Site.Assert.IsTrue(templateList.Length > 1, "GetTemplate operation should return more than one template.");

            #region Capture requirements

            this.VerifyResultOfGetSiteTemplate(getTemplateResult);

            #endregion Capture requirements

            // Invoke the CreateWeb operation without optional parameters, so the return value is expected to contain a url consistent with the expected url.
            // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
            createResult = this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateList[1].Name, 0, false, 0, false, 0, false, true, false, true, false, true, false);
            expectedUrl = expectedUrl.TrimEnd('/');
            string actualUrl = createResult.CreateWeb.Url.TrimEnd('/');

            // If returned value contain a url consistent with the expected url, it means the CreateWeb operation succeed.
            Site.Assert.IsTrue(
                expectedUrl.Equals(actualUrl, StringComparison.CurrentCultureIgnoreCase),
                "Created web's url should be {0}.",
                expectedUrl);

            #region Capture requirements
            // Get a string contains the name and value of the expected properties of the created web.
            string webPropertyDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), webName);
            //// Get each property value by splitting the string.
            Dictionary<string, string> properties = AdapterHelper.DeserializeWebProperties(webPropertyDefault, Constants.ItemSpliter, Constants.KeySpliter);
            uint languageActual = uint.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyLanguage, this.Site)]);
            uint localeActual = uint.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyLocale, this.Site)]);
            string permissionActual = properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyUserNameInPermissions, this.Site)];
            string[] userNameActual = permissionActual.Split(',');
            bool anonymousActual = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyAnonymous, this.Site)]);

            // Get a string contains the name and value of the expected properties of the parent web of the created web.
            string webParentDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), string.Empty);
            //// Get each property value by splitting the string.
            Dictionary<string, string> parentProperties = AdapterHelper.DeserializeWebProperties(webParentDefault, Constants.ItemSpliter, Constants.KeySpliter);
            string parentPermission = parentProperties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyUserNameInPermissions, this.Site)];
            string[] userNameExpected = parentPermission.Split(',');

            // If language is omitted, R515 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R515");

            // Verify MS-SITESS requirement: MS-SITESS_R515
            Site.CaptureRequirementIfAreEqual<uint>(
                language,
                languageActual,
                515,
                @"[In CreateWeb] [language:] If omitted, the subsite to be created MUST use the server’s default language for the user interface.");

            // If locale is omitted, R516 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R516");

            // Verify MS-SITESS requirement: MS-SITESS_R516
            Site.CaptureRequirementIfAreEqual<uint>(
                locale,
                localeActual,
                516,
                @"[In CreateWeb] [locale:] If omitted, specifies that the subsite to be created MUST use the server’s default settings for displaying data.");

            // If uniquePermissions is omitted, R555 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R555, the actual user name is {0}", userNameActual);

            // Verify MS-SITESS requirement: MS-SITESS_R555
            bool isVerifyR555 = AdapterHelper.CompareStringArrays(userNameExpected, userNameActual);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR555,
                555,
                @"[In CreateWeb] [uniquePermissions:] If omitted, the subsite to be created MUST inherit its permissions from its parent site.");

            // If anonymous is omitted, R556 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R556, the actual anonymous is {0}", anonymousActual);

            // Verify MS-SITESS requirement: MS-SITESS_R556
            Site.CaptureRequirementIfIsFalse(
                anonymousActual,
                556,
                @"[In CreateWeb] [anonymous:] If omitted, the anonymous authentication MUST NOT be allowed for the subsite to be created.");

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            this.VerifyOperationCreateWeb();
            #endregion Capture requirements

            // If R3781 is not enabled, that means the CreateWeb operation is not supported, so there is no web to be deleted here.
            if (Common.IsRequirementEnabled(3781, this.Site) && Common.IsRequirementEnabled(3791, this.Site))
            {
                // Invoke the DeleteWeb operation.
                this.sitessAdapter.DeleteWeb(webUrl);

                #region Capture requirements
                // Verify that Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.
                this.VerifyOperationDeleteWeb();
                #endregion Capture requirements
            }
        }

        /// <summary>
        /// This test case is designed to verify GetSiteTemplates and CreateWeb operations when the requested URL is already in use.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC03_CreateWebFailureUrlAlreadyInUse()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site);
            uint getTemplateResult = 0;
            bool isErrorOccured = false;
            SoapException soapException = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // GetTemplate operation should return more than one template.
            Site.Assert.IsTrue(templateList.Length > 1, "GetTemplate operation should return more than one template.");

            #region Capture requirements

            this.VerifyResultOfGetSiteTemplate(getTemplateResult);

            #endregion Capture requirements

            // Try to invoke the CreateWeb operation with url parameter set to the web which is already used (i.e. Subsite1).
            // The first template of the returned template list is a Global template and can't be used to create web ,so templateList[1] is used here.
            try
            {
                this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateList[1].Name, localeId, true, localeId, true, localeId, true, true, true, true, true, true, true);

                Site.Log.Add(LogEntryKind.Comment, "CreateWeb succeed!");
            }
            catch (SoapException ex)
            {
                soapException = ex;
                isErrorOccured = true;
                this.VerifySoapFaultDetail(ex);
            }

            #region Capture requirements
            // When invoke the CreateWeb operation with url parameter set to the web which 
            // is already used (i.e. Subsite1), if an ErrorCode 0x800700b7 is thrown, 
            // it means the expected SOAP fault is thrown, so the following requirements
            // can be captured.
            string errorCode = Common.ExtractErrorCodeFromSoapFault(soapException);

            this.VerifyErrorCodeOfCreateWeb(errorCode);

            // If the error code is 0x800700b7, R282 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R282, the error code is {0}", errorCode);

            // Verify MS-SITESS requirement: MS-SITESS_R282
            Site.CaptureRequirementIfIsTrue(
                isErrorOccured && errorCode.Equals("0x800700b7"),
                282,
                @"[In CreateWeb] If the error code is 0x800700b7, it specifies the location specified by CreateWeb/url is already in use.");

            this.VerifySoapFault(isErrorOccured);

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            this.VerifyOperationCreateWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify GetSiteTemplates and CreateWeb operations when the requested template does not exist.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC04_CreateWebFailureTemplateNotExist()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + this.newSubsite;
            string templateName = string.Empty;
            uint getTemplateResult = 0;
            bool isErrorOccured = false;
            SoapException soapException = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // GetTemplate operation should return more than one template.
            Site.Assert.IsTrue(templateList.Length > 1, "GetTemplate operation should return more than one template.");

            #region Capture requirements

            this.VerifyResultOfGetSiteTemplate(getTemplateResult);

            #endregion Capture requirements

            // If 0 is returned and the templateList is not empty, it means the GetSiteTemplates operation is succeed.
            // Try to invoke the CreateWeb operation with invalid templateName parameter.
            try
            {
                // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
                templateName = templateList[1].Name + "1";
                this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateName, localeId, true, localeId, true, localeId, true, true, true, true, true, true, true);

                Site.Log.Add(LogEntryKind.Comment, "CreateWeb succeed!");
            }
            catch (SoapException ex)
            {
                soapException = ex;
                isErrorOccured = true;
                this.VerifySoapFaultDetail(ex);
            }

            #region Capture requirements

            // When invoke the CreateWeb operation with invalid templateName parameter, if an ErrorCode 0x8102009f is thrown, it means the expected
            // SOAP fault is thrown, so the following requirements can be captured.
            string errorCode = Common.ExtractErrorCodeFromSoapFault(soapException);

            this.VerifyErrorCodeOfCreateWeb(errorCode);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R283, the error code is {0}", errorCode);

            // Verify MS-SITESS requirement: MS-SITESS_R283
            Site.CaptureRequirementIfIsTrue(
                isErrorOccured && errorCode.Equals("0x8102009f"),
                283,
                @"[In CreateWeb] If the error code is 0x8102009f, it specifies the template specified by CreateWeb/templateName does not exist.");

            this.VerifySoapFault(isErrorOccured);

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            this.VerifyOperationCreateWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify DeleteWeb operation when webUrl is non-existent.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC05_DeleteWebFailureNonExistentUrl()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3791, this.Site), @"Test is executed only when R3791Enabled is set to true.");

            #region Variables

            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + Common.GetConfigurationPropertyValue(Constants.NonExistentSiteName, this.Site);
            bool isErrorReturned = false;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            try
            {
                this.sitessAdapter.DeleteWeb(webUrl);

                Site.Log.Add(LogEntryKind.Comment, "DeleteWeb succeed!");
            }
            catch (SoapException)
            {
                isErrorReturned = true;
            }

            #region Capture requirements
            // If the SOAP fault occurs, R411 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R411, the error {0} returned.", isErrorReturned ? "is" : "is not");

            // Verify MS-SITESS requirement: MS-SITESS_R411
            Site.CaptureRequirementIfIsTrue(
                isErrorReturned,
                411,
                @"[In DeleteWeb] [The client sends a DeleteWebSoapIn request message] Otherwise [if delete the subsite unsuccessful], the server MUST return a SOAP fault that is defined in the DeleteWebResponse message.");

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.
            this.VerifyOperationDeleteWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the properties of CreateWeb operation.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC06_CreateWebWithZeroFalse()
        {
            #region Variables
            string subSiteToBeCreated = this.newSubsite;
            string url = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + subSiteToBeCreated;
            string title = Constants.WebTitle;
            string description = Constants.WebDescription;
            string templateName = string.Empty;
            uint language = 0;
            bool languageSpecified = true;
            uint locale = 0;
            bool localeSpecified = true;
            uint collationLocale = 0;
            bool collationLocaleSpecified = true;
            bool uniquePermissions = false;
            bool uniquePermissionsSpecified = true;
            bool anonymous = false;
            bool anonymousSpecified = true;
            bool presence = false;
            bool presenceSpecified = true;
            Template[] templateList;

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            #endregion Variables

            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters.
            this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // GetTemplate operation should return more than one template.
            Site.Assert.IsTrue(templateList.Length > 1, "GetTemplate operation should return more than one template.");

            // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
            templateName = templateList[1].Name;

            // Create web with language, locale, collationLocale zero and uniquePermissions, anonymous, presence false.
            this.sitessAdapter.CreateWeb(url, title, description, templateName, language, languageSpecified, locale, localeSpecified, collationLocale, collationLocaleSpecified, uniquePermissions, uniquePermissionsSpecified, anonymous, anonymousSpecified, presence, presenceSpecified);
            Site.Log.Add(LogEntryKind.Comment, "CreateWeb succeed!");

            #region Capture requirements
            // Get a string contains the name and value of the expected properties of the created web.
            string webPropertyDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), subSiteToBeCreated);
            //// Get each property value by splitting the string.
            Dictionary<string, string> properties = AdapterHelper.DeserializeWebProperties(webPropertyDefault, Constants.ItemSpliter, Constants.KeySpliter);
            uint languageActual = uint.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyLanguage, this.Site)]);
            uint defaultLanguage = uint.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyDefaultLanguage, this.Site)]);
            uint localeActual = uint.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyLocale, this.Site)]);
            uint defaultLocale = uint.Parse(Common.GetConfigurationPropertyValue(Constants.DefaultLCID, this.Site));
            string permissionActual = properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyUserNameInPermissions, this.Site)];
            string[] userNameActual = permissionActual.Split(',');
            bool anonymousActual = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyAnonymous, this.Site)]);

            // Get a string contains the name and value of the expected properties of the parent web of the created web.
            string webParentDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), string.Empty);
            //// Get each property value by splitting the string.
            Dictionary<string, string> parentProperties = AdapterHelper.DeserializeWebProperties(webParentDefault, Constants.ItemSpliter, Constants.KeySpliter);
            string parentPermission = parentProperties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyUserNameInPermissions, this.Site)];
            string[] userNameExpected = parentPermission.Split(',');

            // If language is zero, R552 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R552");

            // Verify MS-SITESS requirement: MS-SITESS_R552
            Site.CaptureRequirementIfAreEqual<uint>(
                defaultLanguage,
                languageActual,
                552,
                @"[In CreateWeb] [language:] If zero, the subsite to be created MUST use the server’s default language for the user interface.");

            // If locale is zero, R553 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R553");

            // Verify MS-SITESS requirement: MS-SITESS_R553
            Site.CaptureRequirementIfAreEqual<uint>(
                defaultLocale,
                localeActual,
                553,
                @"[In CreateWeb] [locale:] If zero, specifies that the subsite to be created MUST use the server’s default settings for displaying data.");

            // If uniquePermissions is false, R519 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R519", userNameActual);

            // Verify MS-SITESS requirement: MS-SITESS_R519
            bool isVerifyR519 = AdapterHelper.CompareStringArrays(userNameExpected, userNameActual);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR519,
                519,
                @"[In CreateWeb] [uniquePermissions:] If set to false, the subsite to be created MUST inherit its permissions from its parent site.");

            // If anonymous is false, R521 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R521", anonymousActual);

            // Verify MS-SITESS requirement: MS-SITESS_R521
            Site.CaptureRequirementIfIsFalse(
                anonymousActual,
                521,
                @"[In CreateWeb] [anonymous:] If set to false, the anonymous authentication MUST NOT be allowed for the subsite to be created.");

            // Verify that Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            this.VerifyOperationCreateWeb();
            #endregion Capture requirements

            // If R3781 is not enabled, that means the CreateWeb operation is not supported, so there is no web to be deleted here.
            if (Common.IsRequirementEnabled(3781, this.Site) && Common.IsRequirementEnabled(3791, this.Site))
            {
                this.sitessAdapter.DeleteWeb(url);
                Site.Log.Add(LogEntryKind.Comment, "DeleteWeb succeed!");

                #region Capture requirements
                //// Verify that Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.
                this.VerifyOperationDeleteWeb();
                #endregion Capture requirements
            }
        }

        /// <summary>
        /// This test case is designed to verify GetSiteTemplates operation.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC07_GetSiteTemplateNotInstalledLCID()
        {
            #region Variables
            uint invalidLocaleId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.NotInstalledLCID, this.Site));
            Template[] templateList;
            uint getTemplateResult = 0;
            string errorCode = string.Empty;
            bool isErrorOccured = false;
            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            try
            {
                // Invoke the GetSiteTemplates operation with invalid parameters, soap fault or empty TemplateList is expected.
                getTemplateResult = this.sitessAdapter.GetSiteTemplates(invalidLocaleId, out templateList);
            }
            catch (SoapException e)
            {
                errorCode = Common.ExtractErrorCodeFromSoapFault(e);
                isErrorOccured = true;
                templateList = null;
            }

            #region Capture requirements
            if (Common.IsRequirementEnabled(400, this.Site))
            {
                // SharePoint Foundation 2010 returns a successful GetSiteTemplatesResponse with an empty TemplateList element
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R400, the getTemplateResult is {0}", getTemplateResult);

                // Verify MS-SITESS requirement: MS-SITESS_R400
                bool templateListEmpty = false;
                if (templateList == null)
                {
                    templateListEmpty = true;
                }
                else
                {
                    templateListEmpty = templateList.Length == 0;
                }

                bool isVerifyR400 = getTemplateResult == 0 && templateListEmpty;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR400,
                    400,
                    @"[In Appendix B: Product Behavior] <7> Section 3.1.4.5.2.2: Implementation does not return the SOAP fault. It returns a successful GetSiteTemplatesResponse with an empty TemplateList element.(Microsoft SharePoint Foundation 2010 and SharePoint Foundation 2013 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1561, this.Site))
            {
                // If the error code is 0x81070209, MS-SITESS_R1561 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R1561, the error code is {0}", errorCode);

                // Verify MS-SITESS requirement: MS-SITESS_R1561
                Site.CaptureRequirementIfIsTrue(
                    isErrorOccured && errorCode.Equals("0x81070209"),
                    1561,
                    @"[In GetSiteTemplatesResponse] [TemplateList:] In case the LCID included in the request message indicates a language that is not installed on the server, Implementation does return a SOAP fault with the error code [0x81070209] specified in the following table.<7> (Windows SharePoint Services 3.0, Office SharePoint Server 2007 and SharePoint Server 2016 follow this behavior.)");
            }
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the successful status of GetSiteTemplates operation.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC08_GetSiteTemplatesSuccessfully()
        {
            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            uint getTemplateResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // If the templateList is not empty, it means the GetSiteTemplates operation is succeed.
            Site.Assert.IsTrue(
                 templateList != null && templateList.Length != 1,
                "GetTemplate operation should return more than one template.");

            #region Capture requirements

            this.VerifyResultOfGetSiteTemplate(getTemplateResult);

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify CreateWeb operations when the presence is false.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC09_CreateWebWithPresenceIsFalse()
        {
            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + this.newSubsite;
            CreateWebResponseCreateWebResult createResult;
            string expectedUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site)
                + "/"
                + this.newSubsite;
            uint getTemplateResult = 0;
            string webName = this.newSubsite;

            #endregion Variables

            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // If the templateList is not empty, it means the GetSiteTemplates operation is succeed.
            Site.Assert.IsTrue(
                 templateList != null && templateList.Length > 1,
                "GetTemplate operation should return more than one template.");

            // Invoke the CreateWeb operation with valid parameters, so the return value is expected to contain a URL consistent with the expected URL.
            // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
            createResult = this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateList[1].Name, localeId, true, localeId, true, localeId, true, true, true, true, true, false, true);
            expectedUrl = expectedUrl.TrimEnd('/');
            string actualUrl = createResult.CreateWeb.Url.TrimEnd('/');
            string webPropertyDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), webName);
            //// Get each property value by splitting the string.
            Dictionary<string, string> properties = AdapterHelper.DeserializeWebProperties(webPropertyDefault, Constants.ItemSpliter, Constants.KeySpliter);
            bool presence = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyPresence, this.Site)]);

            if(Common.IsRequirementEnabled(523001,this.Site))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R523001");

                // Verify MS-SITESS requirement: MS-SITESS_R523001
                Site.CaptureRequirementIfIsTrue(
                    presence,
                    523001,
                    @"[In Appendix B: Product Behavior] Implementation does be enabled for the subsite to be created, when presence set to false, and anonymous authentication is not allowed for the subsite.  <17> Section 3.1.4.9.2.1:  If anonymous authentication is not allowed for the subsite, the online presence information will be enabled on SharePoint Foundation 2010.");
            }

            if (Common.IsRequirementEnabled(523002,this.Site))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R523002");

                // Verify MS-SITESS requirement: MS-SITESS_R523002
                Site.CaptureRequirementIfIsFalse(
                    presence,
                    523002,
                    @"[In Appendix B: Product Behavior]Implementation does not be enabled for the subsite to be created, when presence set to false. (SharePoint Foundation 2013 and above follow this hebavior.)");
            }

            // If R3781 is not enabled, that means the CreateWeb operation is not supported, so there is no web to be deleted here.
            if (Common.IsRequirementEnabled(3781, this.Site) && Common.IsRequirementEnabled(3791, this.Site))
            {
                // Invoke the DeleteWeb operation.
                this.sitessAdapter.DeleteWeb(webUrl);
            }
        }

        /// <summary>
        /// This test case is designed to verify CreateWeb operations when the presence is omitted.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S02_TC10_CreateWebWithPresenceIsOmitted()
        {
            #region Variables

            uint localeId = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templateList;
            string webUrl = Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site)
                + "/"
                + this.newSubsite;
            CreateWebResponseCreateWebResult createResult;
            string expectedUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site)
                + "/"
                + this.newSubsite;
            uint getTemplateResult = 0;
            string webName = this.newSubsite;

            #endregion Variables

            Site.Assume.IsTrue(Common.IsRequirementEnabled(3781, this.Site), @"Test is executed only when R3781Enabled is set to true.");

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetSiteTemplates operation with valid parameters, so result == 0 and templateList.Length > 1 are expected.
            getTemplateResult = this.sitessAdapter.GetSiteTemplates(localeId, out templateList);

            // If the templateList is not empty, it means the GetSiteTemplates operation is succeed.
            Site.Assert.IsTrue(
                 templateList != null && templateList.Length > 1,
                "GetTemplate operation should return more than one template.");

            // Invoke the CreateWeb operation with valid parameters, so the return value is expected to contain a URL consistent with the expected URL.
            // The first template is a Global template and can't be used to create web ,so templateList[1] is used here.
            createResult = this.sitessAdapter.CreateWeb(webUrl, Constants.WebTitle, Constants.WebDescription, templateList[1].Name, localeId, true, localeId, true, localeId, true, true, true, true, true, true, false);
            expectedUrl = expectedUrl.TrimEnd('/');
            string actualUrl = createResult.CreateWeb.Url.TrimEnd('/');
            string webPropertyDefault = this.sutAdapter.GetWebProperties(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site), webName);
            //// Get each property value by splitting the string.
            Dictionary<string, string> properties = AdapterHelper.DeserializeWebProperties(webPropertyDefault, Constants.ItemSpliter, Constants.KeySpliter);
            bool presence = bool.Parse(properties[Common.GetConfigurationPropertyValue(Constants.SubSitePropertyPresence, this.Site)]);

            if (Common.IsRequirementEnabled(527001, this.Site))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R527001");

                // Verify MS-SITESS requirement: MS-SITESS_R527001
                Site.CaptureRequirementIfIsTrue(
                    presence,
                    527001,
                    @"[In Appendix B: Product Behavior] Implementation does be enabled for the subsite to be created, when presence set to omitted, and anonymous authentication is not allowed for the subsite.  <17> Section 3.1.4.9.2.1:  If anonymous authentication is not allowed for the subsite, the online presence information will be enabled on SharePoint Foundation 2010.");
            }

            if (Common.IsRequirementEnabled(527002, this.Site))
            {
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R527002");

                // Verify MS-SITESS requirement: MS-SITESS_R527002
                Site.CaptureRequirementIfIsFalse(
                    presence,
                    527002,
                    @"[In Appendix B: Product Behavior]Implementation does not be enabled for the subsite to be created, when presence set to omitted. (SharePoint Foundation 2013 and above follow this hebavior.)");
            }

            // If R3781 is not enabled, that means the CreateWeb operation is not supported, so there is no web to be deleted here.
            if (Common.IsRequirementEnabled(3781, this.Site) && Common.IsRequirementEnabled(3791, this.Site))
            {
                // Invoke the DeleteWeb operation.
                this.sitessAdapter.DeleteWeb(webUrl);
            }
        }

        /// <summary>
        /// GetSiteTemplatesResult MUST be 0 when the operation succeeded. If actualValue is returned as 0, R153 can be captured.
        /// </summary>
        /// <param name="actualValue">The result of GetSiteTemplates.</param>
        public void VerifyResultOfGetSiteTemplate(uint actualValue)
        {
            // If 0 is returned by actualValue, R153 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R153");

            // Verify MS-SITESS requirement: MS-SITESS_R153
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                actualValue,
                153,
                @"[In GetSiteTemplatesResponse] [GetSiteTemplatesResult:] It MUST be 0 when the operation succeeded.");
        }

        /// <summary>
        /// Protocol server faults using SOAP faults as specified either in [SOAP1.1] section 4.4, or in [SOAP1.2/1] section 5.4.
        /// </summary>
        /// <param name="isSoapFault">Indicate whether the response is a SoapFault.</param>
        public void VerifySoapFault(bool isSoapFault)
        {
            // If an SoapException is thrown, R366 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R366, a SoapFault {0} returned.", isSoapFault ? "is" : "is not");

            // Verify MS-SITESS requirement: MS-SITESS_R366
            Site.CaptureRequirementIfIsTrue(
                isSoapFault,
                366,
                @"[In Transport] Protocol server faults can be returned using SOAP faults as specified either in [SOAP1.1] section 4.4, or in [SOAP1.2/1] section 5.4.");

            // If an SoapException is thrown, R355 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R355, a SoapFault {0} returned.", isSoapFault ? "is" : "is not");

            // Verify MS-SITESS requirement: MS-SITESS_R355
            Site.CaptureRequirementIfIsTrue(
                isSoapFault,
                355,
                @"[In Protocol Details] This protocol [MS-SITESS] allows protocol servers to notify protocol clients of application-level faults using SOAP faults.");

            // If an SoapException is thrown, R356 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R356, a SoapFault {0} returned.", isSoapFault ? "is" : "is not");

            // Verify MS-SITESS requirement: MS-SITESS_R356
            Site.CaptureRequirementIfIsTrue(
                isSoapFault,
                356,
                @"[In Protocol Details] This protocol [MS-SITESS] allows protocol servers to provide additional details for SOAP faults by including either a detail element as specified in [SOAP1.1] section 4.4, or a Detail element as specified in [SOAP1.2/1] section 5.4.5, which conforms to the XML schema of the SOAPFaultDetails complex type specified in section 2.2.4.1.");
        }

        /// <summary>
        /// Verify SoapFaultDetail message.
        /// </summary>
        /// <param name="soapFault">A SoapException contains Soap fault message.</param>
        public void VerifySoapFaultDetail(SoapException soapFault)
        {
            bool isSchemaRight = false;
            string detailBody = SchemaValidation.GetSoapFaultDetailBody(SchemaValidation.LastRawResponseXml.OuterXml);
            ValidationResult detailResult = SchemaValidation.ValidateXml(this.Site, detailBody);
            isSchemaRight = detailResult.Equals(ValidationResult.Success);

            // If the errorSchema is right, R9 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R9");

            // Verify MS-SITESS requirement: MS-SITESS_R9
            Site.CaptureRequirementIfIsTrue(
                isSchemaRight,
                9,
                @"[In SOAPFaultDetails] [The SOAPFaultDetails is defined as:] <s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap"">
              <s:complexType name=""SOAPFaultDetails"">
              <s:sequence>
              <s:element name=""errorstring"" type=""s:string""/>
              <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
              </s:sequence>
              </s:complexType>
              </s:schema>");

            string errorCode = Common.ExtractErrorCodeFromSoapFault(soapFault);

            // If error code is empty, R422 can be verified.
            if (string.IsNullOrEmpty(errorCode))
            {
                // If the errorSchema is right, R422 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R422");

                // Verify MS-SITESS requirement: MS-SITESS_R422
                Site.CaptureRequirementIfIsTrue(
                    isSchemaRight,
                    422,
                    @"[In SOAPFaultDetails] This element [errorcode] is optional.");
            }

            // If error code is not empty, R421 can be verified.
            if (!string.IsNullOrEmpty(errorCode))
            {
                // If the errorSchema is right, R421 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R421, the error code is {0}.", errorCode);

                bool isVerify421 = errorCode.Length.Equals(10) && errorCode.StartsWith("0x", StringComparison.CurrentCulture);

                // Verify MS-SITESS requirement: MS-SITESS_R421
                Site.CaptureRequirementIfIsTrue(
                    isVerify421,
                    421,
                    @"[In SOAPFaultDetails] errorcode: The hexadecimal representation of a four-byte result code.");
            }
        }

        /// <summary>
        /// CreateWeb MUST return a SOAP fault with the specified error code 0x800700b7, 0x8102009f.
        /// </summary>
        /// <param name="actualValue">error code of the SOAP fault</param>
        public void VerifyErrorCodeOfCreateWeb(string actualValue)
        {
            // If error code is 0x800700b7 or 0x8102009f, R281 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R281, the actual value of error code is {0}", actualValue);

            // Verify MS-SITESS requirement: MS-SITESS_R281
            bool isVerifyR281 = actualValue.Equals("0x800700b7") || actualValue.Equals("0x8102009f");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR281,
                281,
                @"[In CreateWeb] If any of the error conditions specified by the following table occur, this method [CreateWeb] MUST return a SOAP fault with the specified error code [0x800700b7, 0x8102009f].");
        }

        /// <summary>
        /// This method is used to verify Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
        /// </summary>
        public void VerifyOperationCreateWeb()
        {
            // If code can run to here, it means Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R3781, Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.");

            // Verify MS-SITESS requirement: MS-SITESS_R3781
            Site.CaptureRequirement(
                3781,
                @"[In Appendix B: Product Behavior] Implementation does support this method [CreateWeb]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// This method is used to verify Microsoft SharePoint Foundation 2010 and above support operation CreateWeb.
        /// </summary>
        public void VerifyOperationDeleteWeb()
        {
            // If code can run to here, it means Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R3791, Microsoft SharePoint Foundation 2010 and above support operation DeleteWeb.");

            // Verify MS-SITESS requirement: MS-SITESS_R3791
            Site.CaptureRequirement(
                3791,
                @"[In Appendix B: Product Behavior] Implementation does support this method [DeleteWeb]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
        }
        #endregion Scenario 2 Manage a subsite

        #endregion Test Cases

        #region Test Case Initialization & Cleanup

        /// <summary>
        /// Test Case Initialization.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.sitessAdapter = Site.GetAdapter<IMS_SITESSAdapter>();
            Common.CheckCommonProperties(this.Site, true);
            this.sutAdapter = Site.GetAdapter<IMS_SITESSSUTControlAdapter>();
            this.newSubsite = "NewSubsite" + Common.FormatCurrentDateTime();
        }

        /// <summary>
        /// Test Case Cleanup.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            // Remove the created subsite for the following products since the CreateWeb and DeleteWeb operations are not supported for SharePointServer2007 and WindowsSharePointServices3.
            if (Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.SharePointServer2010, System.StringComparison.CurrentCultureIgnoreCase)
                || Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.SharePointFoundation2010, System.StringComparison.CurrentCultureIgnoreCase)
                || Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.SharePointServer2013, System.StringComparison.CurrentCultureIgnoreCase)
                || Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.SharePointFoundation2013, System.StringComparison.CurrentCultureIgnoreCase))
            {
                try
                {
                    this.sitessAdapter.DeleteWeb(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site) + "/" + this.newSubsite);
                }
                catch (SoapException e)
                {
                    Site.Log.Add(LogEntryKind.Comment, "S2_ManageSubSite_TestCleanup: ");
                    Site.Log.Add(LogEntryKind.Comment, e.Code.ToString());
                    Site.Log.Add(LogEntryKind.Comment, e.Message);
                }
            }

            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion Test Case Initialization & Cleanup
    }
}