namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System;
    using System.Linq;
    using System.ServiceModel;
    using System.ServiceModel.Channels;
    using System.Text;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class contains all helper methods used in test case level.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// A list of claims provider hierarchy trees.
        /// </summary>
        private static SPProviderHierarchyTree[] allProviders = null;

        /// <summary>
        /// Gets or sets IMS_CPSWSAdapter instance.
        /// </summary>
        protected static IMS_CPSWSAdapter CPSWSAdapter { get; set; }

        /// <summary>
        /// Gets or sets IMS_CPSWSSUTControlAdapter instance.
        /// </summary>
        protected static IMS_CPSWSSUTControlAdapter SutControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets the search string.
        /// </summary>
        protected static string SearchPattern { get; set; }

        #endregion Variables     

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void TestSuiteClassInitialize(TestContext testContext)
        {
            // A method is used to initialize the variables.
            TestClassBase.Initialize(testContext);
            CPSWSAdapter = BaseTestSite.GetAdapter<IMS_CPSWSAdapter>();
            SutControlAdapter = BaseTestSite.GetAdapter<IMS_CPSWSSUTControlAdapter>();
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [ClassCleanup]
        public static void TestSuiteClassCleanup()
        {
            // Cleanup test site, must be called to ensure closing of logs.
            TestClassBase.Cleanup();
        }
        #endregion Test Suite Initialization

        /// <summary>
        /// A method used to get all existed claims providers hierarchy tree in the server.
        /// </summary>
        /// <returns>A return value represents the list of claim providers hierarchy tree.</returns>
        public static SPProviderHierarchyTree[] GetAllProviders()
        {
            SPPrincipalType principalType = SPPrincipalType.All;
            int numberOfLevels = 10;

            if (allProviders == null)
            {
                allProviders = CPSWSAdapter.GetHierarchyAll(null, principalType, numberOfLevels);
            }

            return allProviders;
        }

        /// <summary>
        /// A method used to get claim types by SUT control adapter.
        /// </summary>
        /// <param name="claimProviderNames">A parameter represents a group of provider name.</param>
        /// <returns>A return value represents a list of claim types</returns> 
        public ArrayOfString GetClaimTypesResultBySutAdapter(ArrayOfString claimProviderNames)
        {
            ArrayOfString claimTypesResultBySut = new ArrayOfString();
            string inputClaimProviderNames = InputClaimProviderNames(claimProviderNames);
            string getClaimTypesInSPProviderScript = SutControlAdapter.GetClaimTypesInSPProvider(inputClaimProviderNames);
            if (getClaimTypesInSPProviderScript != null && getClaimTypesInSPProviderScript != string.Empty)
            {
                foreach (string claimType in getClaimTypesInSPProviderScript.Split(','))
                {
                    claimTypesResultBySut.Add(claimType);
                }
            }

            return claimTypesResultBySut;
        }

        /// <summary>
        /// A method used to get claim value types by SUT control.
        /// </summary>
        /// <param name="claimProviderNames">A parameter represents a group of provider name.</param>
        /// <returns>A return value represents a list of claim value types</returns> 
        public ArrayOfString GetClaimValueTypesResultBySutAdapter(ArrayOfString claimProviderNames)
        {
            ArrayOfString claimValueTypesResultBySut = new ArrayOfString();
            string inputClaimProviderNames = InputClaimProviderNames(claimProviderNames);
            string getClaimValueTypesInSPProviderScript = SutControlAdapter.GetClaimValueTypesInSPProvider(inputClaimProviderNames);
            if (getClaimValueTypesInSPProviderScript != null && getClaimValueTypesInSPProviderScript != string.Empty)
            {
                foreach (string claimValueType in getClaimValueTypesInSPProviderScript.Split(','))
                {
                    claimValueTypesResultBySut.Add(claimValueType);
                }
            }

            return claimValueTypesResultBySut;
        }

        /// <summary>
        /// A method used to get entity types by SUT control.
        /// </summary>
        /// <param name="claimProviderNames">A parameter represents a group of provider name.</param>
        /// <returns>A return value represents a list of entity types</returns> 
        public ArrayOfString GetEntityTypesResultBySutAdapter(ArrayOfString claimProviderNames)
        {
            ArrayOfString getEntityTypesResultBySut = new ArrayOfString();
            string inputClaimProviderNames = InputClaimProviderNames(claimProviderNames);
            string getEntityTypesInSPProviderScript = SutControlAdapter.GetEntityTypesInSPProvider(inputClaimProviderNames);
            if (getEntityTypesInSPProviderScript != null && getEntityTypesInSPProviderScript != string.Empty)
            {
                foreach (string entityTypes in getEntityTypesInSPProviderScript.Split(','))
                {
                    getEntityTypesResultBySut.Add(entityTypes);
                }
            }

            return getEntityTypesResultBySut;
        }

        /// <summary>
        /// A method used to generate random GUID.
        /// </summary>
        /// <returns>A return value represents a GUID.</returns>
        public string GenerateGUID()
        {
            Guid guid = new Guid();           
            return guid.ToString();
        }

        /// <summary>
        /// A method used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>A return value represents the random generated string.</returns>
        public string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// A method used to generate an invalid domain user.
        /// </summary>
        /// <returns>A return value represents the invalid domain user.</returns>
        public string GenerateInvalidUser()
        {
            string invalidUser = this.GenerateRandomString(5) + "\"" + this.GenerateRandomString(5);
            return invalidUser;
        }

        /// <summary>
        /// A method used to generate a valid SPClaim of ResolveInput which is used in the ResolveClaim operation.
        /// </summary>
        /// <returns>A return value represents the SPClaim of resolveInput.</returns>
        public SPClaim GenerateSPClaimResolveInput_Valid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();
         
            SPClaim resolveInput = new SPClaim();

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                if (provider.Children.Length != 0)
                {                   
                    resolveInput.ClaimType = provider.Children[0].HierarchyNodeID;
                    resolveInput.Value = this.GenerateRandomString(5);
                    resolveInput.ValueType = this.GenerateGUID();
                    resolveInput.OriginalIssuer = "ClaimProvider:" + provider.ProviderName;                  
                }
            }

            return resolveInput;
        }

        /// <summary>
        /// A method used to generate an invalid SPClaim of ResolveInput.
        /// </summary>
        /// <returns>A return value represents the invalid SPClaim of resolveInput.</returns>
        public SPClaim GenerateSPClaimResolveInput_Invalid()
        {
            SPClaim resolveInput = new SPClaim();
 
            resolveInput.ClaimType = this.GenerateRandomString(10);
            resolveInput.Value = this.GenerateRandomString(5);
            resolveInput.ValueType = this.GenerateGUID();
            resolveInput.OriginalIssuer = "ClaimProvider:" + this.GenerateRandomString(5);
            
            return resolveInput;
        }

        /// <summary>
        /// A method used to do a depth first traverse in a given Provider Hierarchy tree to 
        /// gather provider names and save them in the list of providerNames.
        /// </summary>
        /// <param name="node">A parameter represents a provider hierarchy node.</param>
        /// <param name="providerNames">A parameter represents a list of provider names.</param>
        public void DepthFirstTraverse(SPProviderHierarchyNode node, ref ArrayOfString providerNames)
        {
            providerNames.Add(node.ProviderName);
            if (node.IsLeaf == true)
            {
                return;
            }

            foreach (SPProviderHierarchyNode c in node.Children)
            {
                this.DepthFirstTraverse(c, ref providerNames);
            }
        }

        /// <summary>
        /// A method used to generate a valid SPProviderSearchArguments of IClaimProviderWebService_Search_InputMessage.
        /// </summary>
        /// <returns>A return value represents the valid SPProviderSearchArguments. Return "null" if there is no valid search argument.</returns>
        public SPProviderSearchArguments GenerateProviderSearchArgumentsInput_Valid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();

            SPProviderSearchArguments providerSearchArgumentsInput = null;

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                // Find the first provider tree with at least a child.
                if (provider.Children.Length != 0)
                {
                    providerSearchArgumentsInput = new SPProviderSearchArguments();

                    // Set the search condition by the first child of this tree.
                    providerSearchArgumentsInput.HierarchyNodeID = provider.Children[0].HierarchyNodeID;
                    providerSearchArgumentsInput.ProviderName = provider.Children[0].ProviderName;
                    providerSearchArgumentsInput.MaxCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("MaxCount", this.Site)); 
                    TestSuiteBase.SearchPattern = provider.Children[0].Nm;

                    break;
                }
            }

            return providerSearchArgumentsInput;
        }

        /// <summary>
        /// A method used to generate a valid input condition of SearchAll.
        /// </summary>
        public void GenerateSearchAllInput_Valid()
        {
            // Call the helper method to get all claims providers.
            SPProviderHierarchyTree[] allProviders = TestSuiteBase.GetAllProviders();
                        
            TestSuiteBase.SearchPattern = null;

            foreach (SPProviderHierarchyTree provider in allProviders)
            {
                // Find the first provider tree with at least a child. 
                if (provider.Children.Length != 0)
                {
                    // Get the child provider's name.
                    TestSuiteBase.SearchPattern = provider.Children[0].Nm;
                    break;
                }
            }

            return;
        }

        /// <summary>
        /// A method used to verify argument null exception.
        /// </summary>
        /// <param name="faultException">A parameter represents a fault exception.</param>
        /// <param name="argumentName">A parameter represents an argument name.</param>
        /// <returns>A return value represents the fault exception's type and its argument name.</returns> 
        public bool VerifyArgumentNullException(FaultException faultException, string argumentName)
        {
            MessageFault fault = faultException.CreateMessageFault();
            XmlDocument doc = new XmlDocument();

            doc.Load(fault.GetReaderAtDetailContents());

            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("ex", @"http://schemas.datacontract.org/2004/07/System.ServiceModel");
            XmlNode messageNode = doc.SelectNodes("//ex:Message", xnm).OfType<XmlNode>().SingleOrDefault();
            XmlNode typeNode = doc.SelectNodes("//ex:Type", xnm).OfType<XmlNode>().SingleOrDefault();

            string message = messageNode.InnerText;
            string type = typeNode.InnerText;

            return type.Contains("ArgumentNullException") && message.Contains(argumentName);
        }

        /// <summary>
        /// A method used to verify argument out of range exception.
        /// </summary>
        /// <param name="faultException">A parameter represents a fault exception.</param>
        /// <param name="argumentName">A parameter represents an argument name.</param>
        /// <returns>A return value represents the fault exception's type and its argument name.</returns> 
        public bool VerifyArgumentOutOfRangeException(FaultException faultException, string argumentName)
        {
            MessageFault fault = faultException.CreateMessageFault();
            XmlDocument doc = new XmlDocument();

            doc.Load(fault.GetReaderAtDetailContents());

            XmlNamespaceManager xnm = new XmlNamespaceManager(doc.NameTable);
            xnm.AddNamespace("ex", @"http://schemas.datacontract.org/2004/07/System.ServiceModel");
            XmlNode messageNode = doc.SelectNodes("//ex:Message", xnm).OfType<XmlNode>().SingleOrDefault();
            XmlNode typeNode = doc.SelectNodes("//ex:Type", xnm).OfType<XmlNode>().SingleOrDefault();

            string message = messageNode.InnerText;
            string type = typeNode.InnerText;

            return type.Contains("ArgumentOutOfRangeException") && message.Contains(argumentName);
        }

        /// <summary>
        /// A method used to verify types that are defined by protocol.
        /// </summary>
        /// <param name="resultByProtocol">A parameter represents an expected type list from protocol.</param>
        /// <param name="resultBySutAdapter">A parameter represents an actual type list from SUT adapter.</param>
        /// <returns>Return true value represents the protocol type list is same as script type list, else return false.</returns> 
        public bool VerificationSutResultsAndProResults(ArrayOfString resultByProtocol, ArrayOfString resultBySutAdapter)
        {
            if (resultByProtocol == null)
            {
                throw new ArgumentException("The expected results must not be null or the count must not be 0.");
            }

            if (resultBySutAdapter == null)
            {
                throw new ArgumentException("The actual results must not be null or the count must not be 0.");
            }

            if (resultByProtocol.Count != resultBySutAdapter.Count)
            {
                return false;
            }

            bool retValue = true;
            foreach (string sutItem in resultBySutAdapter)
            {
                if (!resultByProtocol.Any(protocolItem => protocolItem.Equals(sutItem, StringComparison.OrdinalIgnoreCase)))
                {
                    retValue = false;
                    this.Site.Log.Add(LogEntryKind.CheckFailed, "The actual item {0} is not in the protocol results.", sutItem);
                    break;
                }
            }

            if (retValue == true)
            {
                foreach (string protocolItem in resultByProtocol)
                {
                    if (!resultBySutAdapter.Any(sutItem => sutItem.Equals(protocolItem, StringComparison.OrdinalIgnoreCase)))
                    {
                        retValue = false;
                        this.Site.Log.Add(LogEntryKind.CheckFailed, "The expected item {0} is not in the SUT adapter results.", protocolItem);
                        break;
                    }
                }
            }

            return retValue;
        }

        /// <summary>
        ///  A method used to verify remove the duplicated item.
        /// </summary>
        /// <param name="resultByProtocol">A parameter represents an expected type list from protocol.</param>
        /// <returns>Return true value represents the protocol server remove the duplicated item, else return false</returns>
        public bool VierfyRemoveDuplicate(ArrayOfString resultByProtocol)
        {
            for (int i = 0; i < resultByProtocol.Count; i++)
            {
                string claimType = resultByProtocol[i];

                for (int j = i + 1; j < resultByProtocol.Count; j++)
                {
                    if (claimType == resultByProtocol[j])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        [TestInitialize]
        public void TestSuiteBaseInitialize()
        {
            Common.CheckCommonProperties(this.Site, true);

           // Check if MS-CPSWS service is supported in current SUT.
           if (!Common.GetConfigurationPropertyValue<bool>("MS-CPSWS_Supported", this.Site))
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", this.Site);
                this.Site.Assert.Inconclusive("This test suite does not supported under current SUT, because MS-CPSWS_Supported value set to false in MS-CPSWS_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [TestCleanup]
        public void TestSuiteBaseCleanup()
        {
            // Resetting of adapter.
            CPSWSAdapter.Reset();
            SutControlAdapter.Reset();
        }

        /// <summary>
        /// A method used to convert ArrayOfString type to String type.
        /// </summary>
        /// <param name="inputClaimProviderNames">A parameter represents a group of provider name with ArrayOfString type.</param>
        /// <returns>A return value represents a string for provider names, they are separated by commas ',' </returns> 
        private static string InputClaimProviderNames(ArrayOfString inputClaimProviderNames)
        {
            string str = null;
            foreach (string inputClaimProviderName in inputClaimProviderNames)
            {
                str += inputClaimProviderName + ",";
            }

            string results = str.Trim(',');
            return results;
        }
    }
}