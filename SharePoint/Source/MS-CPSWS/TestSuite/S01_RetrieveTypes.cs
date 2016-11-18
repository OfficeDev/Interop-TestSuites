namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
 
    /// <summary>
    /// Scenario 1 Test cases. Test the requirements of 3 operations ClaimTypes, ClaimValueTypes, and EntityTypes. 
    /// These operations are used to retrieve a list of all possible claim types, claim value types and entity types from a list of claim providers available to the protocol client.
    /// </summary>
    [TestClass]
    public class S01_RetrieveTypes : TestSuiteBase
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
        /// A test case used to test ClaimTypes method with providerNames parameter is set to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC01_ClaimTypes_NullProviderNames()
        {
            // Call ClaimTypes method to get claim types with a null providerNames in the request.
            ArrayOfString responseOfClaimTypesResult = CPSWSAdapter.ClaimTypes(null);
            Site.Assert.IsNotNull(responseOfClaimTypesResult, "If the providerNames is NULL, the protocol server MUST use all the available claims providers.");
        }

        /// <summary>
        /// A test case used to test ClaimTypes method with providerNames parameter is set to all of valid provider name.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC02_ClaimTypes_AllValidProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call ClaimTypes method to get claim types with all of valid providerNames in the request.
            ArrayOfString responseOfClaimTypesResult = CPSWSAdapter.ClaimTypes(providerNames);
            Site.Assert.IsNotNull(responseOfClaimTypesResult, "If the providerNames is all of valid providerNames, the protocol server MUST use all the available claims providers.");

            bool isNotDuplicate = VierfyRemoveDuplicate(responseOfClaimTypesResult);
           
            Site.CaptureRequirementIfIsTrue(
                isNotDuplicate,
                146001,
                @"[In ClaimTypes] The protocol server will remove the duplicated claim types from the known basic claim types and the claim providers’ claim types.");

            // Call GetClaimTypesResultBySutAdapter method to get claim types with all of valid providerNames in the request.
            ArrayOfString getClaimTypesResultBySutAdapter = GetClaimTypesResultBySutAdapter(providerNames);
            Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfClaimTypesResult, getClaimTypesResultBySutAdapter), "The claim types returned by the protocol and script should be equal.");

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirement(
                144,
                @"[In ClaimTypes] The protocol server MUST retrieve all known basic claim types. In addition, the protocol server MUST retrieve claim types from claims providers that meet both of the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claim providers are listed in the provider names in the input message.");
        }

        /// <summary>
        /// A test case used to test ClaimTypes method with a valid providerNames parameter.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC03_ClaimTypes_ValidProviderName()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                ArrayOfString providerNames = new ArrayOfString();
                providerNames.Add(provider.ProviderName);

                // Call ClaimTypes method to get claim types with valid providerNames in the request.
                ArrayOfString responseOfClaimTypesResult = CPSWSAdapter.ClaimTypes(providerNames);
                Site.Assert.IsNotNull(responseOfClaimTypesResult, "If the providerNames is a valid providerNames, the protocol server MUST use the current available claims providers.");

                // Call GetClaimTypesResultBySutAdapter method to get claim types with valid providerNames in the request.
                ArrayOfString getClaimTypesResultBySutAdapter = GetClaimTypesResultBySutAdapter(providerNames);
                Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfClaimTypesResult, getClaimTypesResultBySutAdapter), "The claim types returned by the protocol and script should be equal.");
            }
        }

        /// <summary>
        /// A test case used to test ClaimValueTypes method with providerNames parameter is set to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC04_ClaimValueTypes_NullProviderNames()
        {
            // Call ClaimValueTypes method to get claim value types with a null providerNames in the request.
            ArrayOfString responseOfClaimValueTypesResult = CPSWSAdapter.ClaimValueTypes(null);
            Site.Assert.IsNotNull(responseOfClaimValueTypesResult, "If the providerNames is NULL, the protocol server MUST use all the available claims providers.");
        }

        /// <summary>
        /// A test case used to test ClaimValueTypes method with providerNames parameter is set to all of valid provider name.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC05_ClaimValueTypes_AllValidProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call ClaimValueTypes method to get claim value types with all of valid providerNames in the request.
            ArrayOfString responseOfClaimValueTypesResult = CPSWSAdapter.ClaimValueTypes(providerNames);
            Site.Assert.IsNotNull(responseOfClaimValueTypesResult, "If the providerNames is all of valid providerNames, the protocol server MUST use all the available claims providers.");

            bool isNotDuplicatd = VierfyRemoveDuplicate(responseOfClaimValueTypesResult);

            Site.CaptureRequirementIfIsTrue(
               isNotDuplicatd,
               163001,
               @"[In ClaimValueTypes] The protocol server will remove the duplicated claim value types from the known basic claim value types and claims providers' claim value types.");

            // Call GetClaimValueTypesResultBySutAdapter method to get claim value types with all of valid providerNames in the request.
            ArrayOfString getClaimValueTypesResultBySutAdapter = GetClaimValueTypesResultBySutAdapter(providerNames);
            Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfClaimValueTypesResult, getClaimValueTypesResultBySutAdapter), "The claim value types returned by the protocol and script should be equal.");

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirement(
                163,
                @"[In ClaimValueTypes] The protocol server MUST retrieve all known basic claim value types. In addition, the protocol server MUST retrieve claim value types from claims providers that meet both of the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.");
        }

        /// <summary>
        /// A test case used to test ClaimValueTypes method with a valid providerNames parameter.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC06_ClaimValueTypes_ValidProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                ArrayOfString providerNames = new ArrayOfString();
                providerNames.Add(provider.ProviderName);

                // Call ClaimValueTypes method to get claim value types with valid providerNames in the request.
                ArrayOfString responseOfClaimValueTypesResult = CPSWSAdapter.ClaimValueTypes(providerNames);
                Site.Assert.IsNotNull(responseOfClaimValueTypesResult, "If the providerNames is a valid providerNames, the protocol server MUST use the current available claims providers.");

                // Call GetClaimValueTypesResultBySutAdapter method to get claim value types with valid providerNames in the request.
                ArrayOfString getClaimValueTypesResultBySutAdapter = GetClaimValueTypesResultBySutAdapter(providerNames);
                Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfClaimValueTypesResult, getClaimValueTypesResultBySutAdapter), "The claim value types returned by the protocol and script should be equal.");
            }
        }

        /// <summary>
        /// A test case used to test EntityTypes method with providerNames parameter is set to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC07_EntityTypes_NullProviderNames()
        {
            // Call EntityTypes method to get entity types with a null providerNames in the request.
            ArrayOfString responseOfEntityTypesResult = CPSWSAdapter.EntityTypes(null);
            Site.Assert.IsNotNull(responseOfEntityTypesResult, "If the providerNames is NULL, the protocol server MUST use all the available claims providers.");
        }

        /// <summary>
        /// A test case used to test EntityTypes method with providerNames parameter is set to all of valid provider name.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC08_EntityTypes_AllValidProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            bool isEntityTypesSuccess = false;

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call EntityTypes method to get entity types with all of valid providerNames in the request.
            ArrayOfString responseOfEntityTypesResult = CPSWSAdapter.EntityTypes(providerNames);
            Site.Assert.IsNotNull(responseOfEntityTypesResult, "If the providerNames is all of valid providerNames, the protocol server MUST use all the available claims providers.");

            // Call GetEntityTypesResultBySutAdapter method to get entity types with all of valid providerNames in the request.
            ArrayOfString getEntityTypesResultBySutAdapter = GetEntityTypesResultBySutAdapter(providerNames);
            Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfEntityTypesResult, getEntityTypesResultBySutAdapter), "The entity types returned by the protocol and script should be equal.");
            isEntityTypesSuccess = true;

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isEntityTypesSuccess,
                178,
                @"[In EntityTypes] The protocol server MUST retrieve picker entity types from claims providers that meet both of the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.");
        }

        /// <summary>
        /// A test case used to test EntityTypes method with a valid providerNames parameter.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S01_TC09_EntityTypes_ValidProviderNames()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();

            bool isEntityTypesSuccess = false;

            foreach (SPProviderHierarchyTree provider in responseOfGetHierarchyAllResult)
            {
                ArrayOfString providerNames = new ArrayOfString();
                providerNames.Add(provider.ProviderName);

                // Call EntityTypes method to get entity types with valid providerNames in the request.
                ArrayOfString responseOfEntityTypesResult = CPSWSAdapter.EntityTypes(providerNames);
                Site.Assert.IsNotNull(responseOfEntityTypesResult, "If the providerNames is a valid providerNames, the protocol server MUST use the current available claims providers.");
                
                // Call GetEntityTypesResultBySutAdapter method to get entity types with valid providerNames in the request.
                ArrayOfString getEntityTypesResultBySutAdapter = GetEntityTypesResultBySutAdapter(providerNames);
                Site.Assert.IsTrue(this.VerificationSutResultsAndProResults(responseOfEntityTypesResult, getEntityTypesResultBySutAdapter), "The entity types returned by the protocol and script should be equal.");
                isEntityTypesSuccess = true;
            }

            // If the claims providers listed in the provider names in the input message is retrieved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isEntityTypesSuccess,
                178,
                @"[In EntityTypes] The protocol server MUST retrieve picker entity types from claims providers that meet both of the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.");
        }
        #endregion
    }
}