namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System;
    using System.Globalization;
    using System.ServiceModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 2 Test cases. Test the requirements of 2 operations GetHierarchy and GetHierarchyAll. 
    /// These operations are used to retrieve claim provider hierarchy trees from a list of claim providers available to the protocol client.
    /// </summary>
    [TestClass]
    public class S02_RetrieveProviderHierarchyTree : TestSuiteBase
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
        /// A test case used to test GetHierarchy method with hierarchyNodeID parameter is set to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC01_GetHierarchy_NullHierarchyNodeID()
        {
            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            string providerName = null;

            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                // Filter the providers, hierarchy providers which name is started with "_HierarchyProvider_" are not supported by GetHierarchy method.
                if (!provider.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)) && provider.Children.Length != 0)
                {
                    providerName = provider.ProviderName;

                    if (providerName != null)
                    {
                        break;
                    }
                }
            }

            Site.Assume.IsNotNull(providerName, "The providerName should not be null!");

            // Call GetHierarchy method to get a claims provider hierarchy tree with a null hierarchyNodeID in the request.
            SPProviderHierarchyTree responseOfGetHierarchyResult = CPSWSAdapter.GetHierarchy(providerName, principalType, null, numberOfLevels);
            Site.Assert.AreEqual<string>("true", responseOfGetHierarchyResult.IsRoot.ToString().ToLower(CultureInfo.CurrentCulture), "Should return the existing root of current claims provider hierarchy tree.");

            // Verify MS-CPSWS requirement: MS-CPSWS_R209
            Site.CaptureRequirementIfAreEqual<string>(
                "true",
                responseOfGetHierarchyResult.IsRoot.ToString().ToLower(CultureInfo.CurrentCulture),
                209,
                @"[In GetHierarchy] hierarchyNodeID: If NULL is specified, then the protocol server MUST return the existing root of claims provider hierarchy tree.");
        }

        /// <summary>
        /// A test case used to test GetHierarchy method with a valid hierarchyNodeID parameter.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC02_GetHierarchy_ValidHierarchyNodeID()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();
   
            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            string providerName = null;
            string hierarchyNodeID = null;

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                // Filter the providers, hierarchy providers which name is started with "_HierarchyProvider_" are not supported by GetHierarchy method.
                if (provider.ProviderName != null && !provider.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)))
                {
                    providerName = provider.ProviderName;

                    foreach (SPProviderHierarchyNode children in provider.Children)
                    {
                        if (children.HierarchyNodeID != null)
                        {
                            hierarchyNodeID = children.HierarchyNodeID;
                        }
                    }

                    if (providerName != null && hierarchyNodeID != null)
                    {
                        break;
                    }
                }
            }

            Site.Assert.IsNotNull(providerName, "No such claim provider which have a hierarchyNodeID exists in the server!");
            Site.Assert.IsNotNull(hierarchyNodeID, "No such claim provider which have a hierarchyNodeID exists in the server!");

            // Call GetHierarchy method to get a claims provider hierarchy tree with a valid child hierarchyNodeID in the request.
            SPProviderHierarchyTree responseOfGetHierarchyResult = CPSWSAdapter.GetHierarchy(providerName, principalType, hierarchyNodeID, numberOfLevels);
            Site.Assert.AreEqual<string>("false", responseOfGetHierarchyResult.IsRoot.ToString().ToLower(CultureInfo.CurrentCulture), "Should return the hierarchyNodeID specified claims provider hierarchy tree, which is not existing root of current claims provider hierarchy tree.");

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                responseOfGetHierarchyResult.IsRoot.ToString().ToLower(CultureInfo.CurrentCulture) == "false",
                194,
                @"[In GetHierarchy] The protocol server MUST retrieve a claims provider hierarchy tree from the claims provider that meets all the following criteria:
The claims provider name is specified in the input message.
The claims provider is associated with the Web application (1) specified in the input message.
The claims provider supports hierarchy.");
        }

        /// <summary>
        /// A test case used to test GetHierarchy method with a valid numberOfLevels parameter.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC03_GetHierarchy_ValidNumberOfLevels()
        {
            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            string providerName = null;
            bool isGetHierarchySuccess = false;

            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                // Filter the providers, hierarchy providers which name is started with "_HierarchyProvider_" are not supported by GetHierarchy method.
                if (provider.ProviderName != null && !provider.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)))
                {
                    providerName = provider.ProviderName;

                    // Call GetHierarchy method to get a claims provider hierarchy tree with a valid numberOfLevels parameter in the request.
                    SPProviderHierarchyTree responseOfGetHierarchyResult = CPSWSAdapter.GetHierarchy(providerName, principalType, null, numberOfLevels);
                    Site.Assert.IsNotNull(responseOfGetHierarchyResult, "If the numberOfLevels is a valid value, the protocol server MUST use the current available claims providers.");
                    isGetHierarchySuccess = true;

                    // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isGetHierarchySuccess,
                        194,
                        @"[In GetHierarchy] The protocol server MUST retrieve a claims provider hierarchy tree from the claims provider that meets all the following criteria:
The claims provider name is specified in the input message.
The claims provider is associated with the Web application (1) specified in the input message.
The claims provider supports hierarchy.");
                }
            } 
        }

        /// <summary>
        /// A test case used to test GetHierarchy method with numberOfLevels parameter is less than 1.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC04_GetHierarchy_InvalidNumberOfLevels()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            int numberOfLevels = 0;
            string providerName = null;

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                // Filter the providers, hierarchy providers which name is started with "_HierarchyProvider_" are not supported by GetHierarchy method.
                if (provider.ProviderName != null && !provider.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)))
                {
                    providerName = provider.ProviderName;
                    break;
                }
            }

            bool caughtException = false;
            try
            {
                // Call GetHierarchy method with numberOfLevels parameter sets to invalid.
                CPSWSAdapter.GetHierarchy(providerName, principalType, null, numberOfLevels);
            }
            catch (FaultException faultException)
            {
                caughtException = true;

                // If the server returns an ArgumentNullException<""numberOfLevels""> message, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.VerifyArgumentOutOfRangeException(faultException, "numberOfLevels"),
                    583,
                    @"[In GetHierarchy] The protocol server MUST return ArgumentOutOfRangeException<""numberOfLevels""> message if the value of this element [numberOfLevels] is less than 1.");
            }
            finally
            {
                this.Site.Assert.IsTrue(caughtException, "The protocol server should return ArgumentOutOfRangeException<numberOfLevels> message if the value of this element [numberOfLevels] is less than 1.");
            }  
        }

        /// <summary>
        /// A test case used to test GetHierarchyAll method with providerNames parameter is set to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC05_GetHierarchyAll_NullProviderNames()
        {
            SPPrincipalType principalType = SPPrincipalType.User;

            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            // Call GetHierarchyAll method to get a list of claims provider hierarchy trees with a null providerNames in the request.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = CPSWSAdapter.GetHierarchyAll(null, principalType, numberOfLevels);
            Site.Assert.IsNotNull(responseOfGetHierarchyAllResult, "If the providerNames is NULL, the protocol server MUST use all the available claims providers.");
        }

        /// <summary>
        /// A test case used to test GetHierarchyAll method with providerNames parameter is set to all of valid provider name.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC06_GetHierarchyAll_AllOfProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] getAllProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.User;

            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            foreach (SPProviderHierarchyTree provider in getAllProviders)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call GetHierarchyAll method to get a list of claims provider hierarchy trees with all of valid providerNames in the request.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = CPSWSAdapter.GetHierarchyAll(providerNames, principalType, numberOfLevels);
            Site.Assert.IsNotNull(responseOfGetHierarchyAllResult, "If the providerNames is all of valid providerNames, the protocol server MUST use all the available claims providers.");
            
            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                responseOfGetHierarchyAllResult.Length == providerNames.Count,
                219,
                @"[In GetHierarchyAll] The protocol server MUST retrieve claims provider hierarchy trees from claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support hierarchy.");

            Site.CaptureRequirementIfIsTrue(
                responseOfGetHierarchyAllResult.Length == providerNames.Count,
                239,
                @"[In GetHierarchyAllResponse]The protocol server MUST return one claims provider hierarchy tree for each claims provider that match the criteria specified in the input message.");
        }

        /// <summary>
        /// A test case used to test GetHierarchyAll method with valid input parameters.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC07_GetHierarchyAll_ValidInputParameters()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] getAllProviders = TestSuiteBase.GetAllProviders();

            SPPrincipalType principalType = SPPrincipalType.User;
          
            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            foreach (SPProviderHierarchyTree provider in getAllProviders)
            {
                ArrayOfString providerNames = new ArrayOfString();
                providerNames.Add(provider.ProviderName);

                // Call GetHierarchyAll method to get a list of claims provider hierarchy trees with a valid providerNames in the request.
                SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = CPSWSAdapter.GetHierarchyAll(providerNames, principalType, numberOfLevels);
                Site.Assert.IsNotNull(responseOfGetHierarchyAllResult, "If the provider names are valid, the protocol server will return the hierarchy trees that match the claims providers.");
            }
        }

        /// <summary>
        /// A test case used to test GetHierarchyAll method with valid value of numberOfLevels.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC08_GetHierarchyAll_ValidNumberOfLevels()
        {
            // Get the valid numberOfLevels of claims provider hierarchy trees.
            int numberOfLevels = Convert.ToInt32(Common.GetConfigurationPropertyValue("numberOfLevels", Site));

            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            bool isGetHierarchyAllSuccess = false;

            // Call GetHierarchy method to get a claims provider hierarchy tree with a valid numberOfLevels parameter in the request.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = CPSWSAdapter.GetHierarchyAll(null, principalType, numberOfLevels);
            Site.Assert.IsNotNull(responseOfGetHierarchyAllResult, "If the numberOfLevels is a valid value, the protocol server MUST use the current available claims providers.");
            isGetHierarchyAllSuccess = true;

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isGetHierarchyAllSuccess,
                219,
                @"[In GetHierarchyAll] The protocol server MUST retrieve claims provider hierarchy trees from claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support hierarchy.");
        }

        /// <summary>
        /// A test case used to test GetHierarchyAll method with numberOfLevels parameter is less than 1.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S02_TC09_GetHierarchyAll_InvalidNumberOfLevels()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] getAllProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;
            int numberOfLevels = 0;

            foreach (SPProviderHierarchyTree provider in getAllProviders)
            {
                providerNames.Add(provider.ProviderName);
            }

            bool caughtException = false; 
            try
            {
                // Call GetHierarchyAll method with numberOfLevels parameter sets to invalid.
                CPSWSAdapter.GetHierarchyAll(providerNames, principalType, numberOfLevels);
            }
            catch (FaultException faultException)
            {
                caughtException = true; 

                // If the server returns an ArgumentNullException<""numberOfLevels""> message, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.VerifyArgumentOutOfRangeException(faultException, "numberOfLevels"),
                    591,
                    @"[In GetHierarchyAll] The protocol server MUST return an ArgumentOutOfRangeException<""numberOfLevels""> message if the value of this element [numberOfLevels] is less than 1.");
            }
            finally
            {
                this.Site.Assert.IsTrue(caughtException, "The protocol server should return ArgumentOutOfRangeException<numberOfLevels> message if the value of this element [numberOfLevels] is less than 1.");
            }
        }
        #endregion
    }
}