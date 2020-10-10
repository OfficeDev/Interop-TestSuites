namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class contains test case definitions related to GetSite operation.
    /// </summary>
    [TestClass]
    public class S06_GetSite : TestClassBase
    {
        /// <summary>
        /// An instance of protocol adapter class.
        /// </summary>
        private IMS_SITESSAdapter sitessAdapter;

        /// <summary>
        /// An instance of SUT control adapter class.
        /// </summary>
        private IMS_SITESSSUTControlAdapter sutAdapter;

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

        #region Scenario 6 Get information about the site collection

        /// <summary>
        /// This test case is designed to verify the successful status of GetSite.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S06_TC01_GetSiteSucceed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5341, this.Site), @"Test is executed only when R5341Enabled is set to true.");

            #region Variables
            string url = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string getResult = string.Empty;
            Guid expectedGuid = Guid.Empty;
            bool expectedUserCodeEnabled = false;
            Site result;
            Guid siteGuid = Guid.Empty;
            bool siteUserCodeEnabled = false;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Get the site collection identifier of the site collection.
            expectedGuid = new Guid(this.sutAdapter.GetSiteGuid());

            Site.Assert.AreNotEqual<Guid>(
                Guid.Empty,
                expectedGuid,
                "Site's guid should not be null.");

            // Set whether user code is enabled for the site collection. Set user code is true.
            expectedUserCodeEnabled = this.sutAdapter.SetUserCodeEnabled(true);
            Site.Assert.IsTrue(expectedUserCodeEnabled, "The user code should be enabled for the site collection.");

            // Invoke the GetSite operation with valid SiteUrl.
            // getResult is in form: <Site Url=UrlString Id=IdString UserCodeEnabled=UserCodeEnabledString />.
            // If split getResult with '"', the fourth is IdString. And, the sixth is UserCodeEnabledString.
            getResult = this.sitessAdapter.GetSite(url);

            Site.Assert.IsNotNull(getResult, "The GetSiteResult should not be null when invoking the GetSite operation with valid SiteUrl");
            result = AdapterHelper.SiteResultDeserialize(getResult);

            // Get the Id element form result of the succeed GetSite operation.
            siteGuid = new Guid(result.Id);

            // Get the UserCodeEnabled element form result of the succeed GetSite operation.
            bool convertResult = bool.TryParse(result.UserCodeEnabled, out siteUserCodeEnabled);

            #region Capture requirements

            this.VerifyIdString(expectedGuid, siteGuid);

            if (convertResult)
            {
                // The value for UserCodeEnabledString MUST be ""true"" if user code is enabled for the site collection.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R398");

                // Verify MS-SITESS requirement: MS-SITESS_R398
                Site.CaptureRequirementIfAreEqual<bool>(
                    expectedUserCodeEnabled,
                    siteUserCodeEnabled,
                    398,
                    @"[In GetSiteResponse] [GetSiteResult:] The value for UserCodeEnabledString MUST be ""true"" if user code is enabled for the site collection.");
            }
            else
            {
                Site.Assert.Fail("The returned value of the UserCodeEnabled element is not of type bool, the value is : {0}", result.UserCodeEnabled);
            }

            if (Common.IsRequirementEnabled(327001002, this.Site))
            {
                string[] urls = new string[] { result.Url };
                bool[] urlss = this.sitessAdapter.IsScriptSafeUrlUsingCustomizedDomain(urls);

                // If IsScriptSafeUrlUsingCustomizedDomain is true, it indicates a URL is a valid script safe URL for the current site by checking against CustomScriptSafeDomains property of the site collection.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R32702701");

                // Verify MS-SITESS requirement: MS-SITESS_R32702701
                Site.CaptureRequirementIfIsTrue(
                    urlss[0],
                    32702701,
                    @"[InArrayOfBoolean]  boolean: Indicates whether a URL is a valid script safe URL for the current site by checking against CustomScriptSafeDomains property of the site collection.");
            }
            #endregion Capture requirements

            // Set whether user code is enabled for the site collection. Set user code is false.
            expectedUserCodeEnabled = this.sutAdapter.SetUserCodeEnabled(false);
            Site.Assert.IsFalse(expectedUserCodeEnabled, "The user code should not be enabled for the site collection.");

            // Invoke the GetSite operation with valid SiteUrl.
            // getResult is in form: <Site Url=UrlString Id=IdString UserCodeEnabled=UserCodeEnabledString />.
            // If split getResult with '"', the fourth is IdString. And, the sixth is UserCodeEnabledString.
            getResult = this.sitessAdapter.GetSite(url);

            Site.Assert.IsNotNull(getResult, "The GetSiteResult should not be null when invoking the GetSite operation with valid SiteUrl");
            result = AdapterHelper.SiteResultDeserialize(getResult);

            // Get the Id element form result of the succeed GetSite operation.
            siteGuid = new Guid(result.Id);           

            // Get the UserCodeEnabled element form result of the succeed GetSite operation.
            convertResult = bool.TryParse(result.UserCodeEnabled, out siteUserCodeEnabled);

            #region Capture requirements

            this.VerifyIdString(expectedGuid, siteGuid);

            if (convertResult)
            {
                // The value for UserCodeEnabledString MUST be ""false"" if user code is not enabled for the site collection.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R399");

                // Verify MS-SITESS requirement: MS-SITESS_R399
                Site.CaptureRequirementIfAreEqual<bool>(
                    expectedUserCodeEnabled,
                    siteUserCodeEnabled,
                    399,
                    @"[In GetSiteResponse] [GetSiteResult:] The value for UserCodeEnabledString MUST be ""false"" if it is not enabled.");
            }
            else
            {
                Site.Assert.Fail("The returned value of the UserCodeEnabled element is not of type bool, the value is : {0}", result.UserCodeEnabled);
            }

            if (Common.IsRequirementEnabled(326001002, this.Site))
            {
                string[] urls = new string[] { result.Url };
                bool[] urlss = this.sitessAdapter.IsScriptSafeUrl(urls);

                // If IsScriptSafeUrl is false, it indicates the url is not a valid script safe url.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R32602702");

                // Verify MS-SITESS requirement: MS-SITESS_R32602702
                Site.CaptureRequirementIfIsFalse(
                    urlss[0],
                    32602702,
                    @"[InArrayOfBoolean]  boolean: [False] Indicates a URL is not a valid script safe URL for the current site.");

                // If urls is a file full path or URL, R422003 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R422003");

                Uri fileLocation;
                bool isUrl = Uri.TryCreate(urls[0], UriKind.Absolute, out fileLocation);

                // Verify MS-SITESS requirement: MS-SITESS_R422003
                Site.CaptureRequirementIfIsTrue(
                    isUrl,
                    422003,
                    @"[In ArrayOfString] string: A file full path or URL.");
            }
            #endregion Capture requirements

            // If code can run to here, it means that Microsoft SharePoint Foundation 2010 and above support operation GetSite.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5341, Microsoft SharePoint Foundation 2010 and above support operation GetSite.");

            // Verify MS-SITESS requirement: MS-SITESS_R5341
            Site.CaptureRequirement(
                5341,
                @"[In Appendix B: Product Behavior] Implementation does support this method [GetSite]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// This method is used to MS-SITESS requirement: MS-SITESS_R375.
        /// </summary>
        /// <param name="expectedGuid">The expected site collection identifier of the site collection.</param>
        /// <param name="actualGuid">The actual site collection identifier of the site collection.</param>
        public void VerifyIdString(Guid expectedGuid, Guid actualGuid)
        {
            // If the Id element is consistent with the site collection identifier of the site collection, the following requirement can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R375");

            // Verify MS-SITESS requirement: MS-SITESS_R375
            Site.CaptureRequirementIfAreEqual<Guid>(
                expectedGuid,
                actualGuid,
                375,
                @"[In GetSiteResponse] [GetSiteResult:] IdString is a quoted string that is the site collection identifier of the site collection.");
        }

        #endregion Scenario 6 Get informations of site collection

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
        }

        /// <summary>
        /// Test Case Cleanup.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion
    }
}