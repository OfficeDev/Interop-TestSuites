namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class containing test case definitions related to GetUpdatedFormDigest and GetUpdatedFormDigestInformation operations.
    /// </summary>
    [TestClass]
    public class S03_GetUpdatedFormDigest : TestClassBase
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

        #region Scenario 3 GetUpdatedFormDigest

        /// <summary>
        /// This test case is designed to verify the GetUpdatedFormDigest operation.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S03_TC01_GetUpdatedFormDigest()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5361, this.Site), @"Test is executed only when R5361Enabled is set to true.");

            #region Variables
            string currentFormDigest = string.Empty;
            string newFormDigest = string.Empty;
            int formDigestTimeout = int.Parse(Common.GetConfigurationPropertyValue(Constants.ExpireTimePeriodBySecond, this.Site));
            string formDigestValid = null;
            string formDigestExpired = null;
            string formDigestReNewed = null;
            string webPageUrl = Common.GetConfigurationPropertyValue(Constants.WebPageUrl, this.Site);
            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetUpdatedFormDigest operation.
            currentFormDigest = this.sitessAdapter.GetUpdatedFormDigest();

            formDigestValid = this.sutAdapter.PostWebForm(currentFormDigest, webPageUrl);
            Site.Assert.IsTrue(formDigestValid.Contains(Constants.PostWebFormResponse), "digest accept");

            // This sleep is just to wait for security validation returned by the server to expire and add 10 s for buffer, not wait for the server to complete operation or return response.
            Thread.Sleep((1000 * formDigestTimeout) + 10000);

            formDigestExpired = this.sutAdapter.PostWebForm(currentFormDigest, webPageUrl);
            if (Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointFoundation2013 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointFoundation2013SP1 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointServer2013 || 
                Common.GetConfigurationPropertyValue(Constants.SutVersion,this.Site) == Constants.SharePointServer2016 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointServer2019)
            {
                Site.Assert.IsTrue(formDigestExpired.Contains(Constants.TimeOutInformationForSP2013AndSP2016), "digest expired");
            }
            else
            {
                Site.Assert.IsTrue(formDigestExpired.Contains(Constants.TimeOutInformationForSP2007AndSP2010), "digest expired");
            }

            // Invoke the GetUpdatedFormDigest operation again, the returned security validation is expected to be different with the last one.
            newFormDigest = this.sitessAdapter.GetUpdatedFormDigest();

            formDigestReNewed = this.sutAdapter.PostWebForm(newFormDigest, webPageUrl);
            Site.Assert.IsTrue(formDigestReNewed.Contains(Constants.PostWebFormResponse), "New digest accept");

            #region Capture requirements

            // If the validation token changed, the following requirement can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R192");

            // Verify MS-SITESS requirement: MS-SITESS_R192
            Site.CaptureRequirementIfAreNotEqual<string>(
                currentFormDigest,
                newFormDigest,
                192,
                @"[In GetUpdatedFormDigest] In this case [when the client request an updated security validation] the server MUST return a new security validation to the client.");

            // If code can run to here, it means that Microsoft Windows SharePoint Services 3.0 and above support method GetUpdatedFormDigest.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5361, Microsoft Windows SharePoint Services 3.0 and above support method GetUpdatedFormDigest.");

            // Verify MS-SITESS requirement: MS-SITESS_R5361
            Site.CaptureRequirement(
                5361,
                @"[In Appendix B: Product Behavior] <9> Section 3.1.4.6: Implementation does support this operation [GetUpdatedFormDigest].(Windows SharePoint Services 3.0 and above follow this behavior.)");
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the GetUpdatedFormDigestInformation operation.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S03_TC02_GetUpdatedFormDigestInformation()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5381, this.Site), @"Test is executed only when R5381Enabled is set to true.");

            string webPageUrl = Common.GetConfigurationPropertyValue(Constants.WebPageUrl, this.Site);
            FormDigestInformation currentInfo = new FormDigestInformation();
            FormDigestInformation newInfo = new FormDigestInformation();
            string urlStr = string.Empty;
            int formDigestTimeout = int.Parse(Common.GetConfigurationPropertyValue(Constants.ExpireTimePeriodBySecond, this.Site));
            string currentSite = Common.GetConfigurationPropertyValue(Constants.SiteCollectionUrl, this.Site);

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the GetUpdatedFormDigestInformation operation.
            currentInfo = this.sitessAdapter.GetUpdatedFormDigestInformation(urlStr);

            var formDigestValid = this.sutAdapter.PostWebForm(currentInfo.DigestValue, webPageUrl);
            Site.Assert.IsTrue(formDigestValid.Contains(Constants.PostWebFormResponse), "digest accept");

            // If the DigestValue is the valid security validation token, R500 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R500, the DigestValue is {0}.", currentInfo.DigestValue);

            // Verify MS-SITESS requirement: MS-SITESS_R500
            Site.CaptureRequirement(
                500,
                @"[In FormDigestInformation] DigestValue: Security validation token generated by the protocol server.");

            // This sleep is just to wait for security validation returned by the server to expire and add 10 s for buffer, not wait for the server to complete operation or return response.
            Thread.Sleep((1000 * formDigestTimeout) + 10000);

            var formDigestExpired = this.sutAdapter.PostWebForm(currentInfo.DigestValue, webPageUrl);
            if (Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointFoundation2013 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointFoundation2013SP1 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointServer2013 || 
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointServer2016 ||
                Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site) == Constants.SharePointServer2019)
            {
                Site.Assert.IsTrue(formDigestExpired.Contains(Constants.TimeOutInformationForSP2013AndSP2016), "digest expired");
            }
            else
            {
                Site.Assert.IsTrue(formDigestExpired.Contains(Constants.TimeOutInformationForSP2007AndSP2010), "digest expired");
            }

            // If the security validation token do expire after specified timeout seconds, R501 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R501, the TimeoutSeconds is {0}.", currentInfo.TimeoutSeconds);

            // Verify MS-SITESS requirement: MS-SITESS_R501
            Site.CaptureRequirement(
                501,
                @"[In FormDigestInformation] TimeoutSeconds: The time in seconds in which the security validation token will expire after the protocol server generates the security validation token server.");

            // Invoke the GetUpdatedFormDigestInformation operation again, the returned security validation is expected to be different with the last one.
            newInfo = this.sitessAdapter.GetUpdatedFormDigestInformation(urlStr);

            var formDigestReNewed = this.sutAdapter.PostWebForm(newInfo.DigestValue, webPageUrl);
            Site.Assert.IsTrue(formDigestReNewed.Contains(Constants.PostWebFormResponse), "New digest accept");

            FormDigestInformation nullInfo = this.sitessAdapter.GetUpdatedFormDigestInformation(null);

            string expectUrl = currentSite.TrimEnd('/');
            string actualUrl = nullInfo.WebFullUrl.TrimEnd('/');
            bool isVerifyR550 = expectUrl.Equals(actualUrl, StringComparison.CurrentCultureIgnoreCase);

            #region Capture requirements
            // If the url is the current requested site, R550 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R550, the actual URL is {0}.", actualUrl);

            // Verify MS-SITESS requirement: MS-SITESS_R550
            Site.CaptureRequirementIfIsTrue(
                isVerifyR550,
                550,
                @"[In GetUpdatedFormDigestInformation] [url:] If this element is omitted altogether, the protocol server MUST return the FormDigestInformation of the current requested site (2).");
            #endregion Capture requirements

            FormDigestInformation emptyInfo = this.sitessAdapter.GetUpdatedFormDigestInformation(urlStr);

            #region Capture requirements
            expectUrl = currentSite.TrimEnd('/');
            actualUrl = emptyInfo.WebFullUrl.TrimEnd('/');
            bool isVerifyR551 = expectUrl.Equals(actualUrl, StringComparison.CurrentCultureIgnoreCase);

            // If the url is the current requested site, R551 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R551, the actual URL is {0}.", actualUrl);

            // Verify MS-SITESS requirement: MS-SITESS_R551
            Site.CaptureRequirementIfIsTrue(
                isVerifyR551,
                551,
                @"[In GetUpdatedFormDigestInformation] [url:] If this element is included as an empty string, the protocol server MUST return the FormDigestInformation of the current requested site (2).");
            #endregion Capture requirements

            urlStr = Common.GetConfigurationPropertyValue(Constants.SiteCollectionUrl, this.Site);
            FormDigestInformation otherInfo = this.sitessAdapter.GetUpdatedFormDigestInformation(urlStr);

            #region Capture requirements
            expectUrl = urlStr.TrimEnd('/');
            actualUrl = otherInfo.WebFullUrl.TrimEnd('/');
            bool isVerifyR405 = expectUrl.Equals(actualUrl, StringComparison.CurrentCultureIgnoreCase);

            // If the url is a requested site, R405 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R405, the actual URL is {0}.", actualUrl);

            // Verify MS-SITESS requirement: MS-SITESS_R405
            Site.CaptureRequirementIfIsTrue(
                isVerifyR405,
                405,
                @"[In GetUpdatedFormDigestInformation] [url:] Otherwise[If this element is neither omitted altogether nor included as an empty string], the protocol server MUST return the FormDigestInformation of the site that contains the page specified by this element.");

            // If code can run to here, it means that Microsoft SharePoint Foundation 2010 and above support method GetUpdatedFormDigestInformation.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5381, Microsoft SharePoint Foundation 2010 and above support method GetUpdatedFormDigestInformation.");

            // Verify MS-SITESS requirement: MS-SITESS_R5381
            Site.CaptureRequirement(
                5381,
                @"[In Appendix B: Product Behavior] Implementation does support this method [GetUpdatedFormDigestInformation]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            #endregion Capture requirements
        }

        #endregion Scenario 3 GetUpdatedFormDigest

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

            // Since the default value of the form digest timeout is too long to wait, change it to a smaller value.
            int formDigestTimeout = int.Parse(Common.GetConfigurationPropertyValue(Constants.ExpireTimePeriodBySecond, this.Site));
            int temp = this.sutAdapter.SetFormDigestTimeout(formDigestTimeout);
            Site.Assert.AreEqual<int>(formDigestTimeout, temp, "The digest timeout value should be set to the input value.");
        }

        /// <summary>
        /// Test Case Cleanup.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            // Set the form digest timeout value to default value on the server.
            int defaultFormDigestTimeout = int.Parse(Common.GetConfigurationPropertyValue(Constants.DefaultExpireTimePeriod, this.Site));
            int temp = this.sutAdapter.SetFormDigestTimeout(defaultFormDigestTimeout);
            Site.Assert.AreEqual<int>(defaultFormDigestTimeout, temp, "The digest timeout value should be set to the input value.");
            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion
    }
}