//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_CPSWS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The test cases of Scenario 05.
    /// </summary>
    [TestClass]
    public class S05_SearchForEntities : TestSuiteBase
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
        /// A method used to verify the search operation can successfully find the first claims provider tree that contains a child.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S05_TC01_Search()
        {
            // Set principal type for search.
            SPPrincipalType principalType = SPPrincipalType.SecurityGroup;

            // Generate ArrayOfSPProviderSearchArguments.
            List<SPProviderSearchArguments> arrayOfSPProviderSearchArguments = new List<SPProviderSearchArguments>();

            Site.Assume.IsNotNull(this.GenerateProviderSearchArgumentsInput_Valid(), "There should be a valid provider search arguments!");
            SPProviderSearchArguments providerSearchArguments = this.GenerateProviderSearchArgumentsInput_Valid();
            arrayOfSPProviderSearchArguments.Add(providerSearchArguments);

            Site.Assume.IsNotNull(TestSuiteBase.SearchPattern, "The search pattern should not be null!");

            // Call Search operation.
            SPProviderHierarchyTree[] responseOfSearchResult = CPSWSAdapter.Search(arrayOfSPProviderSearchArguments.ToArray(), principalType, TestSuiteBase.SearchPattern);

            Site.Assert.IsNotNull(responseOfSearchResult, "The search result MUST be not null.");
            Site.Assert.IsTrue(responseOfSearchResult.Length >= 1, "The search result MUST contain at least one claims provider.");

            // Get the input provider names list.
            ArrayOfString providerNames = new ArrayOfString();
            providerNames.AddRange(arrayOfSPProviderSearchArguments.Select(root => root.ProviderName));

            // Requirement capture condition.
            bool searchSuccess = false;
            
            foreach (SPProviderHierarchyTree providerTree in responseOfSearchResult)
            {
                if (providerTree.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)))
                {
                    if (providerNames.Contains(providerTree.ProviderName))
                    {
                        searchSuccess = true;
                    }
                    else
                    {
                        // Jump over the Hierarchy Provider tree that the server sees fit to return together with the result Claims provider trees.
                        continue;
                    }
                }               
                else if (providerNames.Contains(providerTree.ProviderName))
                {
                    searchSuccess = true;
                }
                else 
                {
                    Site.Assert.Fail("The provider names in the search result should be contained in the provider names in the input message!");
                }
            }
            
            // Capture requirement 378 by matching the input provider name with the result claims provider name,
            // The search input claims provider already satisfy the condition 1 and 3 in requirement 378 in test environment configuration.
            Site.CaptureRequirementIfIsTrue(
                searchSuccess, 
                378, 
                @"[In Search] The protocol server MUST search across all claims providers that meet all the following criteria:
                The claims providers are associated with the Web application (1) specified in the input message.
                The claims providers are listed in the provider search arguments in the input message.
                The claims providers support search.");

            // Capture requirement 398 by matching searchPattern with the result claims provider's Nm attribute.
            Site.CaptureRequirementIfIsTrue(searchSuccess, 398, @"[In Search] searchPattern: The protocol server MUST search for the string in each claims provider.");

            // Capture requirement 404 by matching searchPattern with the result claims provider's Nm attribute.
            Site.CaptureRequirementIfIsTrue(searchSuccess, 404, @"[In SearchResponse] The protocol server MUST return one claims provider hierarchy tree for each claims provider that contains entities that match the search string.");
        }

        /// <summary>
        /// A method used to verify when searchPattern is null as search input, the server will return an ArgumentNullException (searchPattern) message.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S05_TC02_Search_nullSearchPattern()
        {
            // Set principal type for Search.
            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;

            // Set ArrayOfSPProviderSearchArguments for Search.
            List<SPProviderSearchArguments> arrayOfSPProviderSearchArguments = new List<SPProviderSearchArguments>();

            Site.Assume.IsNotNull(this.GenerateProviderSearchArgumentsInput_Valid(), "There should be a valid provider search arguments!");
            SPProviderSearchArguments providerSearchArguments = this.GenerateProviderSearchArgumentsInput_Valid();
            arrayOfSPProviderSearchArguments.Add(providerSearchArguments);

            bool caughtException = false;
            try
            {
                // Call the Search operation with searchPattern as null.
                CPSWSAdapter.Search(arrayOfSPProviderSearchArguments.ToArray(), principalType, null);
            }
            catch (System.ServiceModel.FaultException faultException)
            {
                caughtException = true;

                // Verify Requirement 652, if the server returns an ArgumentNullException<"searchPattern"> message.
                Site.CaptureRequirementIfIsTrue(this.VerifyArgumentNullException(faultException, "searchPattern"), 652, @"[In Search] If this [searchPattern] is NULL, the protocol server MUST return an ArgumentNullException<""searchPattern""> message.");
            }
            finally
            {
                this.Site.Assert.IsTrue(caughtException, "If searchPattern is NULL, the protocol server should return an ArgumentNullException<searchPattern> message.");
            }
        }

        /// <summary>
        /// A method used to verify the SearchAll operation can successfully find the first claims provider tree that contains a child.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S05_TC03_SearchAll()
        {
            // Get the provider names of all the providers from the hierarchy
            ArrayOfString providerNames = new ArrayOfString();
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();
            providerNames.AddRange(responseOfGetHierarchyAllResult.Select(root => root.ProviderName));
            foreach (SPProviderHierarchyNode node in responseOfGetHierarchyAllResult.SelectMany(root => root.Children))
            {
                this.DepthFirstTraverse(node, ref providerNames);
            }

            // Set search principal Type.
            SPPrincipalType principalType = SPPrincipalType.SecurityGroup;

            // Get the searchPattern string as SearchAll input
            this.GenerateSearchAllInput_Valid();

            // Get max count of matched entities allowed to return for this search.
            int maxCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("MaxCount", Site));

            Site.Assume.IsNotNull(TestSuiteBase.SearchPattern, "The search pattern should not be null!");

            // Search the first claims provider tree which has a child.
            SPProviderHierarchyTree[] responseOfSearchResult = CPSWSAdapter.SearchAll(providerNames, principalType, TestSuiteBase.SearchPattern, maxCount);

            // Requirement capture condition.
            bool searchAllSuccess = false;

            foreach (SPProviderHierarchyTree providerTree in responseOfSearchResult)
            {
                if (providerTree.ProviderName.StartsWith(Common.GetConfigurationPropertyValue("HierarchyProviderPrefix", this.Site)))
                {
                    if (providerNames.Contains(providerTree.ProviderName))
                    {
                        searchAllSuccess = true;
                    }
                    else
                    {
                        // Jump over the Hierarchy Provider tree that the server sees fit to return together with the result Claims provider trees.
                        continue;
                    }
                }
                else if (providerNames.Contains(providerTree.ProviderName))
                {
                    searchAllSuccess = true;
                }
                else
                {
                    Site.Assert.Fail("The provider names in the SearchAll result should be contained in the provider names in the input message!");
                }
            }

            // Capture requirement 417 by matching the input provider name with the result claims provider name,
            // The search input claims provider already satisfy the condition 1 and 3 in requirement 417 in test environment configuration.
            Site.CaptureRequirementIfIsTrue(
                searchAllSuccess, 
                417, 
                @"[In SearchAll] The protocol server MUST search across all claims providers that meet all the following criteria:
                The claims providers are associated with the Web application (1) specified in the input message.
                The claims providers are listed in the provider names in the input message.
                The claims providers support search.");
        }

        /// <summary>
        /// A method used to verify when searchPattern is null as the SearchAll input, the server will return an ArgumentNullException(searchPattern) message.
        /// </summary>
        [TestCategory("MSCPSWS"), TestMethod]
        public void MSCPSWS_S05_TC04_SearchAll_nullSearchPattern()
        {
            // Get the provider names of all the providers from the hierarchy
            ArrayOfString providerNames = new ArrayOfString();
            SPProviderHierarchyTree[] responseOfGetHierarchyAllResult = TestSuiteBase.GetAllProviders();
            providerNames.AddRange(responseOfGetHierarchyAllResult.Select(root => root.ProviderName));
            foreach (SPProviderHierarchyNode node in responseOfGetHierarchyAllResult.SelectMany(root => root.Children))
            {
                this.DepthFirstTraverse(node, ref providerNames);
            }

            // Set search principal Type.
            SPPrincipalType principalType = SPPrincipalType.SharePointGroup;

            // Get max count of matched entities allowed to return for this search.
            int maxCount = Convert.ToInt32(Common.GetConfigurationPropertyValue("MaxCount", Site));

            bool caughtException = false; 
            try
            {
                // Set searchPattern as null and call searchAll.
                CPSWSAdapter.SearchAll(providerNames, principalType, null, maxCount);
            }
            catch (System.ServiceModel.FaultException faultException)
            {
                caughtException = true;

                // Verify Requirement 650, if the server returns an ArgumentNullException<"searchPattern"> message.
                Site.CaptureRequirementIfIsTrue(this.VerifyArgumentNullException(faultException, "searchPattern"), 650, @"[In SearchAll] If this [searchPattern] is NULL, the protocol server MUST return an ArgumentNullException<""searchPattern""> message.");
            }
            finally
            {
                this.Site.Assert.IsTrue(caughtException, "If searchPattern is NULL, the protocol server MUST return an ArgumentNullException<searchPattern> message.");
            }
        }
        #endregion
    }
}
