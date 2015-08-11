namespace Microsoft.Protocols.TestSuites.MS_SHDACCWS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-SHDACCWS.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// A string which indicates the error message template.
        /// </summary>
        protected const string ErrorMessageTemplate = "An error occurred while {0}, the error message is {1}";

        /// <summary>
        /// Gets or sets an instance of interface IMS_SHDACCWSAdapter
        /// </summary>
        protected static IMS_SHDACCWSAdapter SHDACCWSAdapter { get; set; }

        /// <summary>
        /// Gets or sets an instance of interface IMS_SHDACCWSSUTControlAdapter
        /// </summary>
        protected static IMS_SHDACCWSSUTControlAdapter SHDACCWSSUTControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets a random generator using current time seeds.
        /// </summary>
        protected static Random RandomInstance { get; set; }

        /// <summary>
        /// Gets or sets a list type instance used to record all lists added by TestSuiteHelper
        /// </summary>
        protected static string ListNameOfAdded { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the contentType number value on current test case.
        /// </summary>
        protected static uint ListNameCounterOfPerTestCases { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the file name number value on current test case.
        /// </summary>
        protected static uint FileNameCounterOfPerTestCases { get; set; }
 
        #endregion

        #region Test Suite Initialization and clean up

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            // A method is used to initialize the variables.
            TestClassBase.Initialize(testContext);
            
            if (null == SHDACCWSAdapter)
            {
                SHDACCWSAdapter = BaseTestSite.GetAdapter<IMS_SHDACCWSAdapter>();
            }

            if (null == SHDACCWSSUTControlAdapter) 
            {
                SHDACCWSSUTControlAdapter = BaseTestSite.GetAdapter<IMS_SHDACCWSSUTControlAdapter>();
            }
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion Test Suite Initialization and clean up

        #region Test case initialization and clean up
        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {            
            // Check if MS-SHDACCWS service is supported in current SUT.
            if (!Common.GetConfigurationPropertyValue<bool>("MS-SHDACCWS_Supported", this.Site))
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", this.Site);
                this.Site.Assert.Inconclusive("This test suite does not supported under current SUT, because MS-SHDACCWS_Supported value set to false in MS-SHDACCWS_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }

            // Initialize the unique resource counter
            ListNameCounterOfPerTestCases = 0;
            FileNameCounterOfPerTestCases = 0;
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.TestCleanup();
        }

        #endregion Test case initialization and clean up
    }
}