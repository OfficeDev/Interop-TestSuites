namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System;
    using System.ServiceModel;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// Adapter class of MS-CPSWS.
    /// </summary>
    public partial class MS_CPSWSAdapter : ManagedAdapterBase, IMS_CPSWSAdapter
    {
        #region Variables

        /// <summary>
        /// The proxy class.
        /// </summary>
        private ClaimProviderWebServiceClient cpswsClient;

        #endregion Variables

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize(), to set default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into Adapter,Make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-CPSWS";

            // Load Common configuration
            this.LoadCommonConfiguration();

            Common.CheckCommonProperties(this.Site, true);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);

            // Initialize the proxy.
            this.cpswsClient = this.GetClaimProviderWebServiceClient();
             
            // Set Credentials information
            Common.AcceptServerCertificate();
        }
        #endregion

        #region MS-CPSWSAdapter Members

        /// <summary>
        /// A method used to get the claim types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the claim types.</param>
        /// <returns>A return value represents a list of claim types.</returns>
        public ArrayOfString ClaimTypes(ArrayOfString providerNames)
        {
            ArrayOfString responseOfClaimTypes = new ArrayOfString();
            try
            {
                responseOfClaimTypes = this.cpswsClient.ClaimTypes(providerNames);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ClaimTypes] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateClaimTypesResponseData(responseOfClaimTypes);

            return responseOfClaimTypes;
        }

        /// <summary>
        /// A method used to get the claim value types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the claim value types.</param>
        /// <returns>A return value represents a list of claim value types.</returns>
        public ArrayOfString ClaimValueTypes(ArrayOfString providerNames)
        {
            ArrayOfString responseOfClaimValueTypes = new ArrayOfString();
            try
            {
                responseOfClaimValueTypes = this.cpswsClient.ClaimValueTypes(providerNames);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ClaimValueTypes] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateClaimValueTypesResponseData(responseOfClaimValueTypes);

            return responseOfClaimValueTypes;
        }

        /// <summary>
        /// A method used to get the entity types.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names which are used to retrieve the entity types.</param>
        /// <returns>A return value represents a list of entity types.</returns>
        public ArrayOfString EntityTypes(ArrayOfString providerNames)
        {
            ArrayOfString responseOfEntityTypes = new ArrayOfString();
            try
            {
                responseOfEntityTypes = this.cpswsClient.EntityTypes(providerNames);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [EntityTypes] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateEntityTypesResponseData(responseOfEntityTypes);

            return responseOfEntityTypes;
        }

        /// <summary>
        /// A method used to retrieve a claims provider hierarchy tree from a claims provider.
        /// </summary>
        /// <param name="providerName">A parameter represents a provider name.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="hierarchyNodeID">A parameter represents the identifier of the node to be used as root of the returned claims provider hierarchy tree.</param>
        /// <param name="numberOfLevels">A parameter represents the maximum number of levels that can be returned by the protocol server in any of the output claims provider hierarchy tree.</param>
        /// <returns>A return value represents a claims provider hierarchy tree.</returns>
        public SPProviderHierarchyTree GetHierarchy(string providerName, SPPrincipalType principalType, string hierarchyNodeID, int numberOfLevels)
        {
            SPProviderHierarchyTree responseOfGetHierarchy = new SPProviderHierarchyTree();
            try
            {
                responseOfGetHierarchy = this.cpswsClient.GetHierarchy(providerName, principalType, hierarchyNodeID, numberOfLevels);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [GetHierarchy] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateGetHierarchyResponseData(responseOfGetHierarchy);

            return responseOfGetHierarchy;
        }

        /// <summary>
        /// A method used to retrieve a list of claims provider hierarchy trees from a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="numberOfLevels">A parameter represents the maximum number of levels that can be returned by the protocol server in any of the output claims provider hierarchy tree.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        public SPProviderHierarchyTree[] GetHierarchyAll(ArrayOfString providerNames, SPPrincipalType principalType, int numberOfLevels)
        {
            SPProviderHierarchyTree[] responseOfGetHierarchyAll = new SPProviderHierarchyTree[0];
            try
            {
                responseOfGetHierarchyAll = this.cpswsClient.GetHierarchyAll(providerNames, principalType, numberOfLevels);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [GetHierarchyAll] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateGetHierarchyAllResponseData(responseOfGetHierarchyAll);

            return responseOfGetHierarchyAll;
        }

        /// <summary>
        /// A method used to retrieve schema for the current hierarchy provider.
        /// </summary>
        /// <returns>A return value represents the hierarchy claims provider schema.</returns>
        public SPProviderSchema HierarchyProviderSchema()
        {
            SPProviderSchema responseOfHierarchyProviderSchema = new SPProviderSchema();
            try
            {
                responseOfHierarchyProviderSchema = this.cpswsClient.HierarchyProviderSchema();
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [HierarchyProviderSchema] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.VerifyHierarchyProviderSchema(responseOfHierarchyProviderSchema);

            return responseOfHierarchyProviderSchema;
        }

        /// <summary>
        /// A method used to retrieve a list of claims provider schemas from a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <returns>A return value represents a list of claims provider schemas</returns>
        public SPProviderSchema[] ProviderSchemas(ArrayOfString providerNames)
        {
            SPProviderSchema[] responseOfProviderSchemas = new SPProviderSchema[0];
            try
            {
                responseOfProviderSchemas = this.cpswsClient.ProviderSchemas(providerNames);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ProviderSchemas] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.VerifyProviderSchemas(responseOfProviderSchemas);

            return responseOfProviderSchemas;
        }

        /// <summary>
        /// A method used to resolve an input string to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents the input to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        public PickerEntity[] Resolve(ArrayOfString providerNames, SPPrincipalType principalType, string resolveInput)
        {
            PickerEntity[] responseOfResolve = new PickerEntity[0];
            try
            {
                responseOfResolve = this.cpswsClient.Resolve(providerNames, principalType, resolveInput);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [Resolve] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateResolveResponseData(responseOfResolve);

            return responseOfResolve;
        }

        /// <summary>
        /// A method used to resolve an input claim to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result</param>
        /// <param name="resolveInput">A parameter represents the SPClaim to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        public PickerEntity[] ResolveClaim(ArrayOfString providerNames, SPPrincipalType principalType, SPClaim resolveInput)
        {
            PickerEntity[] responseOfResolveClaim = new PickerEntity[0];
            try
            {
                responseOfResolveClaim = this.cpswsClient.ResolveClaim(providerNames, principalType, resolveInput);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ResolveClaim] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateResolveClaimResponseData(responseOfResolveClaim);

            return responseOfResolveClaim;
        }

        /// <summary>
        /// A method used to resolve a list of strings to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents a list of input strings to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        public PickerEntity[] ResolveMultiple(ArrayOfString providerNames, SPPrincipalType principalType, ArrayOfString resolveInput)
        {
            PickerEntity[] responseOfResolveMultiple = new PickerEntity[0];
            try
            {
                responseOfResolveMultiple = this.cpswsClient.ResolveMultiple(providerNames, principalType, resolveInput);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ResolveMultiple] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateResolveMultipleResponseData(responseOfResolveMultiple);

            return responseOfResolveMultiple;
        }

        /// <summary>
        /// A method used to resolve a list of claims to picker entities using a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the result.</param>
        /// <param name="resolveInput">A parameter represents a list of claims to be resolved.</param>
        /// <returns>A return value represents a list of picker entities.</returns>
        public PickerEntity[] ResolveMultipleClaim(ArrayOfString providerNames, SPPrincipalType principalType, SPClaim[] resolveInput)
        {
            PickerEntity[] responseOfResolveMultipleClaim = new PickerEntity[0];
            try
            {
                responseOfResolveMultipleClaim = this.cpswsClient.ResolveMultipleClaim(providerNames, principalType, resolveInput);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [ResolveMultipleClaim] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.ValidateResolveMultipleClaimResponseData(responseOfResolveMultipleClaim);

            return responseOfResolveMultipleClaim;
        }

        /// <summary>
        /// A method used to perform a search for entities on a list of claims providers.
        /// </summary>
        /// <param name="providerSearchArguments">A parameter represents the search arguments.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="searchPattern">A parameter represents the search string.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        public SPProviderHierarchyTree[] Search(SPProviderSearchArguments[] providerSearchArguments, SPPrincipalType principalType, string searchPattern)
        {
            SPProviderHierarchyTree[] responseOfSearch = new SPProviderHierarchyTree[0];
            try
            {
                responseOfSearch = this.cpswsClient.Search(providerSearchArguments, principalType, searchPattern);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [Search] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.VerifySearch(responseOfSearch);

            return responseOfSearch;
        }

        /// <summary>
        /// A method used to perform a search for entities on a list of claims providers.
        /// </summary>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        /// <param name="principalType">A parameter represents which type of picker entities to be included in the output claims provider hierarchy tree.</param>
        /// <param name="searchPattern">A parameter represents the search string.</param>
        /// <param name="maxCount">A parameter represents the maximum number of matched entities to be returned in total across all the output claims provider hierarchy trees.</param>
        /// <returns>A return value represents a list of claims provider hierarchy trees.</returns>
        public SPProviderHierarchyTree[] SearchAll(ArrayOfString providerNames, SPPrincipalType principalType, string searchPattern, int maxCount)
        {
            SPProviderHierarchyTree[] responseOfSearchAll = new SPProviderHierarchyTree[0];
            try
            {
                responseOfSearchAll = this.cpswsClient.SearchAll(providerNames, principalType, searchPattern, maxCount);
                this.CaptureTransportRelatedRequirements();
            }
            catch (FaultException faultEX)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when calling [SearchAll] method:\r\n{0}",
                                faultEX.Message);
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureSOAPFaultRequirement(faultEX);
                throw;
            }

            this.VerifySearchAll(responseOfSearchAll);

            return responseOfSearchAll;
        }

        #endregion

        #region Private helper methods

        /// <summary>
        /// A method used to load Common Configuration
        /// </summary>
        private void LoadCommonConfiguration()
        {
            // Get a specified property value from ptfconfig file.
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);
        }

        /// <summary>
        /// A method used to get the claim provider web service client
        /// </summary>
        /// <returns>A return value represents the claim provider web service client.</returns>
        private ClaimProviderWebServiceClient GetClaimProviderWebServiceClient()
        {
            TransportProtocol currentTransportValue = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            string endpointName = string.Empty;
            string targetAddressValue = string.Empty;
            switch (currentTransportValue)
            {
                case TransportProtocol.HTTP:
                    {
                        endpointName = Common.GetConfigurationPropertyValue("HttpEndPointName", this.Site);
                        targetAddressValue = Common.GetConfigurationPropertyValue("TargetHTTPServiceUrl", this.Site);

                        break;
                    }

                case TransportProtocol.HTTPS:
                    {
                        endpointName = Common.GetConfigurationPropertyValue("HttpsEndPointName", this.Site);
                        targetAddressValue = Common.GetConfigurationPropertyValue("TargetHTTPSServiceUrl", this.Site);
                        break;
                    }
            }

            EndpointAddress targetAddress = new EndpointAddress(targetAddressValue);
            ClaimProviderWebServiceClient protocolClient = WcfClientFactory.CreateClient<ClaimProviderWebServiceClient, IClaimProviderWebService>(this.Site, endpointName, targetAddress);
            
            // Setting credential 
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            protocolClient.ClientCredentials.Windows.ClientCredential.UserName = userName;
            protocolClient.ClientCredentials.Windows.ClientCredential.Password = password;
            protocolClient.ClientCredentials.Windows.ClientCredential.Domain = domain;
            protocolClient.ClientCredentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
            protocolClient.Open();

            return protocolClient;
        }

        #endregion
    }
}