namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test cases of Scenario 03.
    /// </summary>
    [TestClass]
    public class S03_RetrieveProviderSchema : TestSuiteBase
    {
        #region Test suite initialization and cleanup
        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        /// <param name="testContext">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.TestSuiteClassInitialize(testContext);
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // Cleanup test site, must be called to ensure closing of logs.
            TestSuiteBase.TestSuiteClassCleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// A method used to verify retrieving the schema of the current hierarchy provider successfully.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S03_TC01_HierarchyProviderSchema()
        {
            // Retrieve the schema of the current hierarchy provider.
            SPProviderSchema responseOfHierarchyProviderSchemaResult = CPSWSAdapter.HierarchyProviderSchema();
            if (responseOfHierarchyProviderSchemaResult == null)
            {
                Site.Assert.Inconclusive("There is no schema for the current hierarchy provider in the test environment!", responseOfHierarchyProviderSchemaResult);
            }
        }

        /// <summary>
        /// A method used to verify retrieving the schemas of the specific claims providers successfully.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S03_TC02_ProviderSchemas()
        {
            // Get the provider names of all the providers from the hierarchy
            ArrayOfString providerNames = new ArrayOfString();
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();
            providerNames.AddRange(responseOfGetHierarchyAllResult.Select(root => root.ProviderName));
            foreach (SPProviderHierarchyNode node in responseOfGetHierarchyAllResult.SelectMany(root => root.Children))
            {
                this.DepthFirstTraverse(node, ref providerNames);
            }
            
            // Get schemas of the claims providers specified in the list of provider names in the request.
            SPProviderSchema[] responseOfProviderSchemaResult = CPSWSAdapter.ProviderSchemas(providerNames);
            Site.Assert.IsNotNull(responseOfProviderSchemaResult, "The schemas of the specific claims providers should not be null!");
        }

        /// <summary>
        /// A method used to verify retrieving schemas of all the available claims providers by not specify provider names in the request.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S03_TC03_ProviderSchemas_NoInputProviderNames()
        {
            // Get schemas of all the available claims providers by not specify provider names in the request.
            SPProviderSchema[] responseOfProviderSchemaResult = CPSWSAdapter.ProviderSchemas(null);
            Site.Assert.IsNotNull(responseOfProviderSchemaResult, "The schemas of all the available claims providers should not be null!");
        }
        #endregion
    }
}