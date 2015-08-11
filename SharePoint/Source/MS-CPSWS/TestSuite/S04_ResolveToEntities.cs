namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System.ServiceModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 4 Test cases. Test resolve related operations and requirements,
    /// include resolving input strings/claims to picker entities using a list of claim providers.
    /// </summary>
    [TestClass]
    public class S04_ResolveToEntities : TestSuiteBase
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
        /// A test case used to test resolve method with valid input string.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC01_ResolveString()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.SecurityGroup;
            string resolveInput = string.Empty;
            bool isResolveSuccess = false;

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                if (provider.EntityData.Length != 0)
                {
                    providerNames.Add(provider.ProviderName);

                    foreach (PickerEntity entityData in provider.EntityData)
                    {
                        resolveInput = entityData.DisplayText;

                        // Call Resolve method to resolve an input string to picker entities using a list of claims providers.
                        PickerEntity[] responseOfResolveResult = CPSWSAdapter.Resolve(providerNames, principalType, resolveInput);                     
                        Site.Assert.IsNotNull(responseOfResolveResult, "Resolve result should not null.");
                        isResolveSuccess = true;
                    }
                }
            }

            // If the claims providers listed in the provider names in the input message is resolved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResolveSuccess,
                280,
                @"[In Resolve] The protocol server MUST resolve across all claims providers that meet all the following criteria:
The claims providers are associated with the Web application specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support resolve.");
        }

        /// <summary>
        /// This test case is used to test typical resolve claim scenario.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC02_ResolveClaim_Valid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.SecurityGroup;
            SPClaim resolveInput = GenerateSPClaimResolveInput_Valid();
            bool isResolveClaimSuccess = false;

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                if (provider.Children.Length != 0)
                {
                    providerNames.Add(provider.ProviderName);                    
                }

                // Call Resolve claim method to resolve an SPClaim to picker entities using a list of claims providers.
                PickerEntity[] responseOfResolveClaimResult = CPSWSAdapter.ResolveClaim(providerNames, principalType, resolveInput);
                Site.Assert.IsNotNull(responseOfResolveClaimResult, "The resolve claim result should not be null.");
                isResolveClaimSuccess = true;               
            }

            // If the claims providers listed in the provider names in the input message is resolved successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isResolveClaimSuccess,
                303,
                @"[In ResolveClaim] The protocol server MUST resolve across all claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support resolve.");
        }

        /// <summary>
        /// This test case is used to resolve 2 valid users to picker entities.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC03_ResolveMultipleStrings_AllValid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.User;
            ArrayOfString resolveInput = new ArrayOfString();
            resolveInput.Add(Common.GetConfigurationPropertyValue("OwnerLogin", this.Site));
            resolveInput.Add(Common.GetConfigurationPropertyValue("ValidUser", this.Site));

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call ResolveMultiple method to resolve 2 valid users to picker entities using a list of claims providers.
            PickerEntity[] responseOfResolveMultipleResult = CPSWSAdapter.ResolveMultiple(providerNames, principalType, resolveInput);
            Site.Assert.IsNotNull(responseOfResolveMultipleResult, "The resolve multiple result should not be null.");

            // If the resolve multiple result contains 2 picker entities, that is to say, one picker entity in the response corresponding to one user in the request, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleResult.Length,
                2,
                345,
                @"[In ResolveMultipleResponse] The list [ResolveMultipleResult] MUST contain one and only one picker entity per string in the input.");

            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleResult.Length,
                2,
                325,
                @"[In ResolveMultiple] The protocol server MUST resolve across all claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support resolve.");
        }

        /// <summary>
        /// This test case is used to resolve 2 users to picker entities, one is valid and another is invalid.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC04_ResolveMultipleStrings_SomeValid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.User;
            ArrayOfString resolveInput = new ArrayOfString();
            resolveInput.Add(Common.GetConfigurationPropertyValue("OwnerLogin", this.Site));
            resolveInput.Add(this.GenerateInvalidUser());

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                providerNames.Add(provider.ProviderName);
            }

            // Call Resolve multiple method to resolve 2 users to picker entities, one valid and another invalid.
            PickerEntity[] responseOfResolveMultipleResult = CPSWSAdapter.ResolveMultiple(providerNames, principalType, resolveInput);
            Site.Assert.IsNotNull(responseOfResolveMultipleResult, "The resolve multiple result should not be null.");

            // If the resolve multiple result contains 2 picker entities, that is to say, one picker entity in the response corresponding to one user in the request, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleResult.Length,
                2,
                345,
                @"[In ResolveMultipleResponse] The list [ResolveMultipleResult] MUST contain one and only one picker entity per string in the input.");

            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleResult.Length,
                2,
                325,
                @"[In ResolveMultiple] The protocol server MUST resolve across all claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support resolve.");
        }

        /// <summary>
        /// This test case is used test resolve multiple method with resolveInput parameter sets to null.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC05_ResolveMultiple_NullResolveInput()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.User;          

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                providerNames.Add(provider.ProviderName);
            }

            bool caughtException = false;
            try
            {
                // Call Resolve multiple method with resolveInput parameter sets to null.
                CPSWSAdapter.ResolveMultiple(providerNames, principalType, null);
            }
            catch (FaultException faultException)
            {
                caughtException = true;

                // If the server returns an ArgumentNullException<""resolveInput""> message, then the following requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.VerifyArgumentNullException(faultException, "resolveInput"),
                    626,
                    @"[In ResolveMultiple] If this [resolveInput] is NULL, the protocol server MUST return an ArgumentNullException<""resolveInput""> message.");
            }
            finally
            {
                this.Site.Assert.IsTrue(caughtException, "If resolveInput is NULL, the protocol server should return an ArgumentNullException<resolveInput> message.");
            }
        }

        /// <summary>
        /// This test case is used test resolve 2 claims to picker entities, one valid and another invalid.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S04_TC06_ResolveMultipleClaim_SomeValid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            ArrayOfString providerNames = new ArrayOfString();
            SPPrincipalType principalType = SPPrincipalType.SecurityGroup;
            SPClaim[] resolveInput = new SPClaim[2];

            resolveInput[0] = this.GenerateSPClaimResolveInput_Valid();
            resolveInput[1] = this.GenerateSPClaimResolveInput_Invalid();

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                if (provider.Children.Length != 0)
                {
                    providerNames.Add(provider.ProviderName);                             
                }
            }

            // Call Resolve multiple claim method to resolve 2 claims to picker entities using a list of claims providers.
            PickerEntity[] responseOfResolveMultipleClaimResult = CPSWSAdapter.ResolveMultipleClaim(providerNames, principalType, resolveInput);
            Site.Assert.IsNotNull(responseOfResolveMultipleClaimResult, "The resolve multiple claim result should not be null.");

            // If the resolve multiple result contains 2 picker entities, that is to say, one picker entity in the response corresponding to one user in the request, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleClaimResult.Length,
                2,
                369,
                @"[In ResolveMultipleClaimResponse] ResolveMultipleClaimResult: The list MUST contain one and only one picker entity per one claim in the input.");

            Site.CaptureRequirementIfAreEqual<int>(
                responseOfResolveMultipleClaimResult.Length,
                2,
                350,
                @"[In ResolveMultipleClaim] The protocol server MUST resolve across all claims providers that meet all the following criteria:
The claims providers are associated with the Web application (1) specified in the input message.
The claims providers are listed in the provider names in the input message.
The claims providers support resolve.");
        }
        #endregion
    }
}