namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Request;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// MS-ASDOC protocol adapter.
        /// </summary>
        private IMS_ASDOCAdapter asdocAdapter;

        /// <summary>
        /// Gets MS-ASDOC protocol adapter.
        /// </summary>
        protected IMS_ASDOCAdapter ASDOCAdapter
        {
            get { return this.asdocAdapter; }
        }

        #endregion

        #region Test case initialize and cleanup

        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.asdocAdapter = Site.GetAdapter<IMS_ASDOCAdapter>();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #endregion

        #region Test case base method

        /// <summary>
        /// Get document class value of a specific folder or document.
        /// </summary>
        /// <param name="linkId">UNC path of shared folder or document in server.</param>
        /// <returns>Search command response.</returns>
        protected SearchResponse SearchCommand(string linkId)
        {
            // Initialize a store object
            SearchStore store = new SearchStore { Name = SearchName.DocumentLibrary.ToString(), Query = new queryType() };

            if (linkId != null)
            {
                // Give the query values
                queryTypeEqualTo subQuery = new queryTypeEqualTo { LinkId = string.Empty, Value = linkId };
                store.Query.ItemsElementName = new ItemsChoiceType2[] { ItemsChoiceType2.EqualTo };
                store.Query.Items = new object[] { subQuery };
            }

            store.Options = new Options1 { ItemsElementName = new ItemsChoiceType6[2] };
            store.Options.ItemsElementName[0] = ItemsChoiceType6.UserName;
            store.Options.ItemsElementName[1] = ItemsChoiceType6.Password;
            store.Options.Items = new string[2];
            store.Options.Items[0] = Common.GetConfigurationPropertyValue("UserName", this.Site);
            store.Options.Items[1] = Common.GetConfigurationPropertyValue("UserPassword", this.Site);

            // Create a search command request.
            SearchRequest searchRequest = Common.CreateSearchRequest(new SearchStore[] { store });

            // Get search command response.
            SearchResponse searchResponse = this.ASDOCAdapter.Search(searchRequest);
            Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, searchResponse.StatusCode, "The call should be successful.");

            return searchResponse;
        }

        #endregion
    }
}