namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with CellSubRequest operation.
    /// </summary>
    [TestClass]
    public class MS_WOPI_S01_Cell : S01_Cell
    {
        #region Variables

        /// <summary>
        /// Gets or sets a IMS_WOPISUTControlAdapter type instance.
        /// </summary>
        protected static IMS_WOPISUTControlAdapter SutController { get; set; }

        /// <summary>
        /// Gets or sets a IMS_WOPISUTManageCodeControlAdapter type instance.
        /// </summary>
        protected static IMS_WOPIManagedCodeSUTControlAdapter WopiSutManageCodeControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets a CurrentTestClientName instance.
        /// </summary>
        protected static string CurrentTestClientName { get; set; }

        #endregion 

        /// <summary>
        /// Execute the share test cases' initialization.
        /// </summary>
        /// <param name="testContext">A parameter represents the test context.</param>
        [ClassInitialize]
        public static void MSWOPISharedTestClassInitialize(TestContext testContext)
        {
            // Execute the MS-FSSHTTP test cases' initialization
            SharedTestSuiteBase.Initialize(testContext);

            // Execute the MS-WOPI initialization
            if (!TestSuiteHelper.VerifyRunInSupportProducts(MS_WOPI_S01_Cell.BaseTestSite))
            {
                return;
            }

            TestSuiteHelper.InitializeTestSuite(MS_WOPI_S01_Cell.BaseTestSite);
            SutController = TestSuiteHelper.WOPISutControladapter;
            WopiSutManageCodeControlAdapter = TestSuiteHelper.WOPIManagedCodeSUTControlAdapter;
            CurrentTestClientName = TestSuiteHelper.CurrentTestClientName;
        }

        /// <summary>
        /// Execute the share test cases' cleanup.
        /// </summary>
        [ClassCleanup]
        public static void MSWOPISharedTestClassCleanup()
        {   
            // Execute the MS-WOPI clean up.
            TestSuiteHelper.CleanUpDiscoveryProcess(CurrentTestClientName, SutController);

            // Execute the SharedTest class level's clean up
            SharedTestSuiteBase.Cleanup();
        }

        /// <summary>
        /// Execute the test case level initialization.
        /// </summary>
        [TestInitialize]
        public void TestCaseLevelInitializeMethod()
        {
            TestSuiteHelper.PerformSupportProductsCheck(this.Site);
            TestSuiteHelper.PerformSupportCobaltCheck(this.Site);
        }

        /// <summary>
        /// This method is used to get WOPI token and add headers.
        /// </summary>
        /// <param name="requestFileUrl">A parameter represents the file URL.</param>
        /// <param name="userName">A parameter represents the user name we used.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain.</param>
        protected override void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            // Get WOPI token and add headers for the file exists.
            TestSuiteHelper.InitializeContextForShare(requestFileUrl, userName, password, domain, CellStoreOperationType.NormalCellStore, this.Site);
        }

        /// <summary>
        /// This method is used to merge the configuration of the WOPI and FSSHTTP.
        /// </summary>
        /// <param name="site">A parameter represents the site.</param>
        protected override void MergeConfigurationFile(ITestSite site)
        {
            // Merge configuration for share test case.
            TestSuiteHelper.MergeConfigurationFileForShare(this.Site);
        }
    }
}