namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Contain test cases designed to test [MS-ONESTORE] protocol.
    /// </summary>
    [TestClass]
    public partial class TestSuiteBase : TestClassBase
    {
        #region Variables
         /// <summary>
        /// Gets or sets the shared Adapter instance.
        /// </summary>
        protected IMS_FSSHTTP_FSSHTTPBAdapter SharedAdapter { get; set; }

        /// <summary>
        /// Gets or sets the Adapter instance.
        /// </summary>
        protected IMS_ONESTOREAdapter Adapter { get; set; }
        /// <summary>
        /// Gets or sets the userName.
        /// </summary>
        protected string UserName { get; set; }

        /// <summary>
        /// Gets or sets the password for the user specified in the UserName property.
        /// </summary>
        protected string Password { get; set; }

        /// <summary>
        /// Gets or sets the domain.
        /// </summary>
        protected string Domain { get; set; }
        /// <summary>
        /// A string value represents the protocol short name for the shared test cases, it is used in runtime. If plan to run the shared test cases, the WOPI server must implement the MS-FSSHTTP.
        /// </summary>
        private const string SharedTestCasesProtocolShortName = "MS-FSSHTTP-FSSHTTPB";
        /// <summary>
        /// A string value represents the protocol short name for the MS-ONESTORE.
        /// </summary>
        private const string OneStoreProtocolShortName = "MS-ONESTORE";

        /// <summary>
        /// A value indicate performing the merge PTF configuration file once.
        /// </summary>
        private static bool isPerformMergeOperation;
        #endregion Variables

        #region Test Case Initialization
        /// <summary>
        /// Initialize the test.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            if (!isPerformMergeOperation)
            {
                this.Site.DefaultProtocolDocShortName = SharedTestCasesProtocolShortName;
                // Get the name of common configuration file.
                string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
                // Merge the common configuration.
                Common.MergeGlobalConfig(commonConfigFileName, this.Site);
                Common.MergeSHOULDMAYConfig(this.Site);
                this.Site.DefaultProtocolDocShortName = OneStoreProtocolShortName;
                Common.MergeSHOULDMAYConfig(this.Site);
                isPerformMergeOperation = true;
            }
            this.SharedAdapter = Site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
            this.Adapter = Site.GetAdapter<IMS_ONESTOREAdapter>();
            this.UserName = Common.GetConfigurationPropertyValue("UserName1", this.Site);
            this.Password = Common.GetConfigurationPropertyValue("Password1", this.Site);
            this.Domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
        }

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #endregion Test Case Initialization
	}
}
