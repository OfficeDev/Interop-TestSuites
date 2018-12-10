namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Contain test cases designed to test [MS-ONESTORE] protocol.
    /// </summary>
    [TestClass]
    public partial class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets or sets the Adapter instance.
        /// </summary>
        protected IMS_FSSHTTP_FSSHTTPBAdapter Adapter { get; set; }

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
        #endregion Variables

        #region Test Case Initialization

        /// <summary>
        /// Initialize the test.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.Site.DefaultProtocolDocShortName = SharedTestCasesProtocolShortName;
            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            // Merge the common configuration.
            Common.MergeGlobalConfig(commonConfigFileName, this.Site);
            Common.MergeSHOULDMAYConfig(this.Site);
            this.Site.DefaultProtocolDocShortName = OneStoreProtocolShortName;
            Common.MergeSHOULDMAYConfig(this.Site);
            this.Adapter = Site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
            this.UserName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            this.Password = Common.GetConfigurationPropertyValue("Password", this.Site);
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
        /// <summary>
        /// initialize the shared context based on the specified request file URL, user name, password and domain.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        protected void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            SharedContext context = SharedContext.Current;

            if (string.Equals("HTTP", Common.GetConfigurationPropertyValue("TransportType", this.Site), System.StringComparison.OrdinalIgnoreCase))
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpTargetServiceUrl", this.Site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpEndPointName", this.Site);
            }
            else
            {
                context.TargetUrl = Common.GetConfigurationPropertyValue("HttpsTargetServiceUrl", this.Site);
                context.EndpointConfigurationName = Common.GetConfigurationPropertyValue("HttpsEndPointName", this.Site);
            }
            context.Site = this.Site;
            context.OperationType = OperationType.FSSHTTPCellStorageRequest;
            context.UserName = userName;
            context.Password = password;
            context.Domain = domain;
        }

        /// <summary>
        /// A method used to create a CellRequest object and initialize it.
        /// </summary>
        /// <returns>A return value represents the CellRequest object.</returns>
        protected FsshttpbCellRequest CreateFsshttpbCellRequest()
        {
            FsshttpbCellRequest cellRequest = new FsshttpbCellRequest();

            // MUST be great or equal to OxFA12994 
            cellRequest.Version = 0xFA12994;

            // MUST be 12 
            cellRequest.ProtocolVersion = 12;

            // MUST be 11 
            cellRequest.MinimumVersion = 11;

            // MUST be 0x9B069439F329CF9C 
            cellRequest.Signature = 0x9B069439F329CF9C;

            // Set the user agent GUID. 
            cellRequest.GUID = FsshttpbCellRequest.UserAgentGuid;

            // Set the value which MUST be 1. 
            cellRequest.RequestHashingSchema = new Compact64bitInt(1u);
            return cellRequest;
        }

        /// <summary>
        /// A method used to create a QueryChanges CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents subRequest id.</param>
        /// <param name="reserved">A parameter that must be set to zero.</param>
        /// <param name="isAllowFragments">A parameter represents that if to allow fragments.</param>
        /// <param name="isExcludeObjectData">A parameter represents if to exclude object data.</param>
        /// <param name="isIncludeFilteredOutDataElementsInKnowledge">A parameter represents if to include the serial numbers of filtered out data elements in the response knowledge.</param>
        /// <param name="reserved1">A parameter represents a 4-bit reserved field that must be set to zero.</param>
        /// <param name="isStorageManifestIncluded">A parameter represents if to include the storage manifest.</param>
        /// <param name="isCellChangesIncluded">A parameter represents if to include the cell changes.</param>
        /// <param name="reserved2">A parameter represents a 6-bit reserved field that must be set to zero.</param>
        /// <param name="cellId">A parameter represents if the Query Changes are scoped to a specific cell. If the Cell ID is 0x0000, no scoping restriction is specified.</param>
        /// <param name="maxDataElements">A parameter represents the maximum data elements to return.</param>
        /// <param name="queryChangesFilterList">A parameter represents how the results of the query will be filtered before it is returned to the client.</param>
        /// <param name="knowledge">A parameter represents what the client knows about a state of a file.</param>
        /// <returns>A return value represents QueryChanges CellSubRequest object.</returns>
        protected QueryChangesCellSubRequest BuildFsshttpbQueryChangesSubRequest(
                                ulong subRequestId,
                                int reserved = 0,
                                bool isAllowFragments = false,
                                bool isExcludeObjectData = false,
                                bool isIncludeFilteredOutDataElementsInKnowledge = true,
                                int reserved1 = 0,
                                bool isStorageManifestIncluded = true,
                                bool isCellChangesIncluded = true,
                                int reserved2 = 0,
                                CellID cellId = null,
                                ulong? maxDataElements = null,
                                List<Filter> queryChangesFilterList = null,
                                Knowledge knowledge = null)
        {
            QueryChangesCellSubRequest queryChange = new QueryChangesCellSubRequest(subRequestId);

            queryChange.Reserved = reserved;
            queryChange.AllowFragments = Convert.ToInt32(isAllowFragments);
            queryChange.ExcludeObjectData = Convert.ToInt32(isExcludeObjectData);
            queryChange.IncludeFilteredOutDataElementsInKnowledge = Convert.ToInt32(isIncludeFilteredOutDataElementsInKnowledge);
            queryChange.Reserved1 = reserved1;

            queryChange.IncludeStorageManifest = Convert.ToInt32(isStorageManifestIncluded);
            queryChange.IncludeCellChanges = Convert.ToInt32(isCellChangesIncluded);
            queryChange.Reserved2 = reserved2;

            if (cellId == null)
            {
                cellId = new CellID(new ExGuid(0, Guid.Empty), new ExGuid(0, Guid.Empty));
            }

            queryChange.CellId = cellId;

            if (maxDataElements != null)
            {
                queryChange.MaxDataElements = new Compact64bitInt(maxDataElements.Value);
            }

            queryChange.QueryChangeFilters = queryChangesFilterList;
            queryChange.Knowledge = knowledge;

            return queryChange;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        protected CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content)
        {
            return this.CreateCellSubRequest(requestToken, base64Content, Convert.FromBase64String(base64Content).Length);
        }

        /// <summary>
        /// A method used to create a CellSubRequest object and initialize it.
        /// </summary>
        /// <param name="requestToken">A parameter represents Request token.</param>
        /// <param name="base64Content">A parameter represents serialized subRequest.</param>
        /// <param name="binaryDataSize">A parameter represents the number of bytes of data in the SubRequestData element of a cell sub-request.</param>
        /// <returns>A return value represents CellSubRequest object.</returns>
        protected CellSubRequestType CreateCellSubRequest(ulong requestToken, string base64Content, long binaryDataSize)
        {
            CellSubRequestType cellRequestType = new CellSubRequestType();
            cellRequestType.SubRequestToken = requestToken.ToString();
            CellSubRequestDataType subRequestData = new CellSubRequestDataType();
            subRequestData.BinaryDataSize = binaryDataSize;
            subRequestData.Text = new string[1];
            subRequestData.Text[0] = base64Content;

            cellRequestType.SubRequestData = subRequestData;

            return cellRequestType;
        }

        /// <summary>
        /// A method used to create a CellSubRequest object for QueryChanges and initialize it.
        /// </summary>
        /// <param name="subRequestId">A parameter represents the subRequest identifier.</param>
        /// <returns>A return value represents the CellRequest object for QueryChanges.</returns>
        protected CellSubRequestType CreateCellSubRequestEmbeddedQueryChanges(ulong subRequestId)
        {
            FsshttpbCellRequest cellRequest = this.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = this.BuildFsshttpbQueryChangesSubRequest(subRequestId);
            cellRequest.AddSubRequest(queryChange, null);

            CellSubRequestType cellSubRequest = this.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            return cellSubRequest;
        }
    }
}
