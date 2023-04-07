namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with query change operation.
    /// </summary>
    [TestClass]
    public abstract class S12_QueryChanges : SharedTestSuiteBase
    {

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="context">VSTS test context.</param>
        [ClassInitialize]
        public static void TestClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void TestClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion 

        #region Test Case Initialization

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S12_QueryChangesInitialization()
        {
            // Initialize the default file URL.
            this.DefaultFileUrl = this.PrepareFile();

        }

        #endregion

        #region Test Case

        #region DataElementTypeFilter Exclude
        /// <summary>
        /// This test method aims to verify DataElementTypeFilter exclude none data element type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC01_QueryChanges_DataElementTypeFilter_ExcludeNone()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType subRequestWithoutFilter = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse cellStorageResponseWithoutFilter = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestWithoutFilter });
            CellSubResponseType subResponseWithoutFilter = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponseWithoutFilter, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseWithoutFilter.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponseWithoutFilter = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponseWithoutFilter, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponseWithoutFilter, this.Site);

            DataElementTypeFilter noneFilterType = new DataElementTypeFilter(DataElementType.None);
            noneFilterType.FilterOperation = 0;
            List<Filter> filters = new List<Filter>();
            filters.Add(noneFilterType);

            // Create query changes request with the specified filters.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            this.Site.Assert.AreEqual<int>(
                    queryResponseWithoutFilter.DataElementPackage.DataElements.Count,
                    queryResponse.DataElementPackage.DataElements.Count,
                    "Exclude None will not exclude any data elements.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R10003
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         10003,
                         @"[In Appendix B: Product Behavior] Implementation support Query Changes sub-request with filters in the current test suite. (Microsoft SharePoint Server 2013 follow this behavior.)");
            }
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter exclude StorageManifestDataElement type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC02_QueryChanges_DataElementTypeFilter_ExcludeStorageManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.StorageManifestDataElementData);
            filterType.FilterOperation = 0;
            List<Filter> filters = new List<Filter>();
            filters.Add(filterType);

            // Create query changes request with the specified filters.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            bool isExcludeStorageManifestDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.StorageManifestDataElementData) == null;
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isExcludeStorageManifestDataElement,
                "The StorageManifestDataElement should be excluded from the response.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter exclude CellManifestDataElement type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC03_QueryChanges_DataElementTypeFilter_ExcludeCellManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.CellManifestDataElementData);
            filterType.FilterOperation = 0;
            List<Filter> filters = new List<Filter>();
            filters.Add(filterType);

            // Create query changes request with the specified filters.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            bool isExcludeCellManifestDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData) == null;
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isExcludeCellManifestDataElement,
                "The CellManifestDataElement should be excluded from the response.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter exclude RevisionManifestDataElement type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC04_QueryChanges_DataElementTypeFilter_ExcludeRevisionManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.RevisionManifestDataElementData);
            filterType.FilterOperation = 0;
            List<Filter> filters = new List<Filter>();
            filters.Add(filterType);

            // Create query changes request with the specified filters.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            bool isExcludeRevisionManifestDataElementData = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.RevisionManifestDataElementData) == null;
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isExcludeRevisionManifestDataElementData,
                "The RevisionManifestDataElementData should be excluded from the response.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter exclude ObjectGroupDataElement type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC05_QueryChanges_DataElementTypeFilter_ExcludeObjectGroupDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.ObjectGroupDataElementData);
            filterType.FilterOperation = 0;
            List<Filter> filters = new List<Filter>();
            filters.Add(filterType);

            // Create query changes request with the specified filters.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            bool isExcludeObjectGroupDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.ObjectGroupDataElementData) == null;
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isExcludeObjectGroupDataElement,
                "The ObjectGroupDataElement should be excluded from the response.");
        }

        #endregion

        #region DataElementTypeFilter Include

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter include none StorageIndexDataElementData type.  
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC06_QueryChanges_DataElementTypeFilter_IncludeStorageIndexDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.StorageIndexDataElementData);
            filterType.FilterOperation = 1;
            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 2, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            bool isIncludeStorageIndexDataElement = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementType == DataElementType.StorageIndexDataElementData);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isIncludeStorageIndexDataElement,
                "The response should only contains the StorageIndexDataElement.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter include none StorageManifestDataElement type. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC07_QueryChanges_DataElementTypeFilter_IncludeStorageManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.StorageManifestDataElementData);
            filterType.FilterOperation = 1;
            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 2, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            bool isIncludeStorageManifestDataElement = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementType == DataElementType.StorageManifestDataElementData);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isIncludeStorageManifestDataElement,
                "The response should only contains the StorageManifestDataElement.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter include none CellManifestDataElement type. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC08_QueryChanges_DataElementTypeFilter_IncludeCellManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.CellManifestDataElementData);
            filterType.FilterOperation = 1;
            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, false, 0, true, true, 0, null, 2, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            bool isIncludeCellManifestDataElement = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            this.Site.Assert.IsTrue(
                isIncludeCellManifestDataElement,
                "The response should only contains the CellManifestDataElement.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter include none RevisionManifestDataElement type. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC09_QueryChanges_DataElementTypeFilter_IncludeRevisionManifestDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.RevisionManifestDataElementData);
            filterType.FilterOperation = 1;
            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 2, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            bool isIncludeRevisionManifestDataElement = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementType == DataElementType.RevisionManifestDataElementData);

            this.Site.Assert.IsTrue(
                isIncludeRevisionManifestDataElement,
                "The response should only contains the RevisionManifestDataElement.");
        }

        /// <summary>
        /// This test method aims to verify DataElementTypeFilter include none ObjectGroupDataElement type. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC10_QueryChanges_DataElementTypeFilter_IncludeObjectGroupDataElement()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            DataElementTypeFilter filterType = new DataElementTypeFilter(DataElementType.ObjectGroupDataElementData);
            filterType.FilterOperation = 1;
            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 2, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            bool isIncludeObjectGroupDataElement = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementType == DataElementType.ObjectGroupDataElementData);

            this.Site.Assert.IsTrue(
                isIncludeObjectGroupDataElement,
                "The response should only contains the ObjectGroupDataElement.");
        }

        #endregion

        #region DataElementIDs
        /// <summary>
        /// This test method aims to verify DataElementIDsFilter filter include one specified data element ID.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC11_QueryChanges_DataElementIDsFilter_Include()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 100, null, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            // Get the first element id.
            ExGuid e1 = queryResponse.DataElementPackage.DataElements[0].DataElementExtendedGUID;

            List<ExGuid> extendedGuid = new List<ExGuid>();
            extendedGuid.Add(e1);
            ExGUIDArray extendedArray = new ExGUIDArray(extendedGuid);
            DataElementIDsFilter filterType = new DataElementIDsFilter(extendedArray);
            filterType.FilterOperation = 1;
            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;

            List<Filter> filters = new List<Filter>();
            filters.Add(allFilter);
            filters.Add(filterType);

            // Query change with the data element id.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, 100, filters, null);
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            bool isOnlyReturnSpecifiedId = queryResponse.DataElementPackage.DataElements.All(dataElement => dataElement.DataElementExtendedGUID.Equals(e1));

            this.Site.Assert.IsTrue(
                isOnlyReturnSpecifiedId,
                "The server only responses the data element with the specified extended guid {0}",
                e1.GUID);
        }

        #endregion

        #region File Size and Type
        /// <summary>
        /// This test method aims to verify requirements when the file is larger than 1MB.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC12_QueryChanges_BigFile()
        {
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);

            // Initialize the service
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
        }

        /// <summary>
        /// This test method aims to verify requirements when the file is zip file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC13_QueryChanges_ZipFile()
        {
            string fileUrl = Common.GetConfigurationPropertyValue("ZipFile", this.Site);

            // Initialize the service
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
        }

        /// <summary>
        /// This test method aims to verify requirements when the file is one note file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC14_QueryChanges_OneNoteFile()
        {
            string url = Common.GetConfigurationPropertyValue("OneNoteFile", Site);
            this.InitializeContext(url, this.UserName01, this.Password01, this.Domain);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentSerialNumber());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(url, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
        }
        #endregion

        #region Flag Attribute
        /// <summary>
        /// This method is used to test query changes with the allow fragment flag is true.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC15_QueryChanges_AllowFragments_One()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1348, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Allow Fragments 2 flag.");
            }
            
            // Initialize the service
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments flag with the value true.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, 10000, null, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            Site.Log.Add(
                LogEntryKind.Debug,
                "If the client specifies anything higher [than limit of DoS mitigation], the server truncates it. Actually it {0}",
                queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().PartialResult ? "truncates" : "does not truncate");
            
            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R444
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the Partial Result flag is set to true, then capture MS-FSSHTTPB_R444.
                Site.CaptureRequirementIfIsTrue(
                         queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().PartialResult,
                         "MS-FSSHTTPB",
                         444,
                         @"[In Query Changes] Maximum Data Elements (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies, in bytes, the limit of data elements at which the server starts breaking up the results into partial results.");

                // For the requirement MS-FSSHTTPB_R990351, it is not fully validated, because it cost too much to validate its size.
                Site.CaptureRequirementIfIsTrue(
                         queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().PartialResult,
                         "MS-FSSHTTPB",
                         990351,
                         @"[In Query Changes] Maximim Data Elements (variable): If the client specifies anything higher [than limit of DoS mitigation], the server truncates it to this value [Max Data Elements].");
            }
            else
            {
                this.Site.Assert.IsTrue(queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().PartialResult, "[In Query Changes] Max Data Elements (variable): A compact unsigned 64-bit integer (section 2.2.1.1) that specifies limit of DoS mitigation at which the server starts breaking up the results into partial results.");
            }

            DataElement fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsNotNull(
                         fragDataElement,
                         "MS-FSSHTTPB",
                         432,
                         @"[In Query Changes] B - Allow Fragments (1 bit): If set, a bit that specifies to allow fragments.");
            }
            else
            {
                this.Site.Assert.IsNotNull(
                    fragDataElement,
                    @"[In Query Changes] B - Allow Fragments (1 bit): If set, a bit that specifies to allow fragments.");
            }
        }

        /// <summary>
        /// This method is used to test query changes with the allow fragment flag is false.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC16_QueryChanges_AllowFragments_Zero()
        {
            // Initialize the service
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments flag with the value true.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, null, 100, null, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            DataElement fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);
        }

        /// <summary>
        /// The method uses to verify whether the Serial Numbers of filtered out data elements is included in the response Knowledge when D - Include Filtered Out Data Elements In Knowledge is set or not.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC17_QueryChanges_IncludeFilteredOutDataElementsInKnowledge()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Query Changes sub-request with filters in the current test suite.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Send request with no filter and get the GUID array.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Build all Filters and Hierarchy Filter
            ExGuid exguid = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.StorageIndexDataElementData).DataElementExtendedGUID;
            byte[] bytes = new byte[21];
            exguid.GUID.ToByteArray().CopyTo(bytes, 1);

            // These two bytes are magic number which is not documented in the open specification in the current stage, but it will be wrote down in the future.
            bytes[0] = 0x29;
            bytes[17] = 0x01;
            HierarchyFilter filterType1 = new HierarchyFilter(bytes);
            filterType1.FilterOperation = 1;
            filterType1.Depth = HierarchyFilterDepth.Deep;

            List<Filter> filters = new List<Filter>();
            AllFilter allFilter = new AllFilter();
            allFilter.FilterOperation = 0;
            filters.Add(allFilter);
            filters.Add(filterType1);

            // Send request with all Filters, Hierarchy Filter and the flag "Include Filtered Out Data Elements In Knowledge" as false.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChangeSubRequest = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, false, 0, true, true, 0, null, null, filters, null);
            cellRequest.AddSubRequest(queryChangeSubRequest, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);

            // Get the fsshttpbResponse which include Knowledge element
            FsshttpbResponse queryResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            // Send request with all Filters, Hierarchy Filter and the flag "Include Filtered Out Data Elements In Knowledge" as true.
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChangeSubRequestSecond = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, true, true, 0, true, true, 0, null, null, filters, null);
            cellRequestSecond.AddSubRequest(queryChangeSubRequestSecond, null);
            CellSubRequestType cellSubRequestSecond = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            CellStorageResponse cellStorageResponseSecond = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequestSecond });
            CellSubResponseType subResponseSecond = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponseSecond, 0, 0, this.Site);

            // Get the fsshttpbResponse which include Knowledge element
            FsshttpbResponse queryResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponseSecond, this.Site);

            bool isVerified = queryResponseFirst.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges.Count
                    < queryResponseSecond.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges.Count;

            Site.Log.Add(
                LogEntryKind.Debug,
                "When Include Filtered Out Data Elements In Knowledge is set, the server responds the specialized knowledge number is {0}",
                queryResponseSecond.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges.Count);

            Site.Log.Add(
                LogEntryKind.Debug,
                "When Include Filtered Out Data Elements In Knowledge is not set, the server responds the specialized knowledge number is {0}",
                queryResponseFirst.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges.Count);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Expect the specialized knowledge number is larger when the Include Filtered Out Data Elements In Knowledge is set, and actually it {0}",
                isVerified ? "is" : "is not");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2148, MS-FSSHTTPB_R4113
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerified,
                         "MS-FSSHTTPB",
                         2148,
                         @"[In Query Changes] D - Include Filtered Out Data Elements In Knowledge (1 bit): If set, a bit that specifies to include the Serial Numbers (section 2.2.1.9) of filtered out data elements in the response Knowledge (section 2.2.1.13).");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4113, SharedContext.Current.Site))
                {
                    Site.CaptureRequirementIfIsTrue(
                             isVerified,
                             "MS-FSSHTTPB",
                             4113,
                             @"[In Appendix B: Product Behavior] If D - Include Filtered Out Data Elements In Knowledge is not set, the Serial Numbers of filtered out data elements are not included in the response Knowledge. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
                }
            }
            else
            {
                Site.Assert.IsTrue(isVerified, "Include Filtered Out Data Elements In Knowledge (1 bit): If set, a bit that specifies to include the serial numbers (section 2.2.1.9) of filtered out data elements in the response knowledge.");
            }
        }

        /// <summary>
        /// The method uses to verify whether the storage manifest is included when F - Include Storage Manifest (1 bit) is set.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC18_QueryChanges_IncludeStorageManifest_One()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 437, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Storage Manifest flag.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with Include Storage Manifest with the value true.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, false, 0, true, true, 0, null, null, null, null);
            cellRequest.AddSubRequest(queryChange, null);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            bool isIncludeStorageManifest = fsshttpbResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.StorageManifestDataElementData) != null;

            Site.Log.Add(
                LogEntryKind.Debug,
                "When Include Storage Manifest (1 bit) is set, the server responds the storage manifest data element, actually it {0}",
                isIncludeStorageManifest ? "does" : "does not");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R437
            // If the storage manifest is returned, then capture MS-FSSHTTPB_R437
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isIncludeStorageManifest,
                         "MS-FSSHTTPB",
                         437,
                         @"[In Query Changes] F - Include Storage Manifest (1 bit): If set, a bit that specifies to include the Storage Manifest. (Microsoft SharePoint Server 2013/Microsoft SharePoint Foundation 2013 follow this behavior.)");
            }
            else
            {
                Site.Assert.IsTrue(isIncludeStorageManifest, @"[In Query Changes] F - Include Storage Manifest (1 bit): If set, a bit that specifies to include the storage manifest. (Microsoft SharePoint Server 2013/Microsoft SharePoint Foundation 2013 follow this behavior.)");
            }
        }

        /// <summary>
        /// The method uses to verify whether the storage manifest is included when F - Include Storage Manifest (1 bit) is not set.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC19_QueryChanges_IncludeStorageManifest_Zero()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4115, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Storage Manifest flag.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with Include Storage Manifest with the value false.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChanges = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, false, 0, false, true, 0, null, null, null, null);
            cellRequest.AddSubRequest(queryChanges, null);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            bool notIncludeStorageManifest = fsshttpbResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.StorageManifestDataElementData) == null;

            Site.Log.Add(
                LogEntryKind.Debug,
                "When include Storage Manifest is not set, the storage manifest is not included. But actually it {0}",
                notIncludeStorageManifest ? "is not include" : "is included");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4115
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         notIncludeStorageManifest,
                         "MS-FSSHTTPB",
                         4115,
                         @"[In Appendix B: Product Behavior] If F - Include Storage Manifest is not set, the Storage Manifest is not included. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.IsTrue(
                    notIncludeStorageManifest,
                    @"[In Query Changes] F - C Include Storage Manifest (1 bit): otherwise[If D - C Include Storage Manifest is not set], the storage manifest is not included. (Microsoft SharePoint Server 2013/Microsoft SharePoint Foundation 2013 follow this behavior.)");
            }
        }

        /// <summary>
        /// The method uses to verify whether the cell changes is included when G - Include Cell Changes is set or not.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC20_QueryChanges_IncludeCellChanges()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4117, this.Site) && !Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 438, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Cell Changes flag.");
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 438, this.Site))
            {
                // Initialize the service
                this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

                // Query changes from the protocol server with G - Include Cell Changes as true
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, false, 0, true, true, 0, null, null, null, null);
                cellRequest.AddSubRequest(queryChange, null);
                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "The first PutChanges operation should succeed.");
                FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
                bool isIncludeCellChanges = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData) != null;

                Site.Log.Add(
                LogEntryKind.Debug,
                "When Include Cell Changes is set, the cell manifest is included. Actually it {0}",
                isIncludeCellChanges ? "is included" : "is not included");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R438
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsTrue(
                             isIncludeCellChanges,
                             "MS-FSSHTTPB",
                             438,
                             @"[In Query Changes] G - Include Cell Changes (1 bit): If set, a bit that specifies to include cell changes. (Microsoft SharePoint Server 2013/Microsoft SharePoint Foundation 2013 follow this behavior.)");
                }
                else
                {
                    this.Site.Assert.IsTrue(isIncludeCellChanges, "[In Query Changes] G - Include Cell Changes (1 bit): If set, a bit that specifies to include cell changes.");
                }
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4117, this.Site))
            {
                // Initialize the service
                this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

                // Query changes from the protocol server with G - Include Cell Changes as false
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                QueryChangesCellSubRequest queryChangeNotInclude = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, false, 0, true, false, 0, null, null, null, null);
                cellRequest.AddSubRequest(queryChangeNotInclude, null);
                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "The first PutChanges operation should succeed.");
                FsshttpbResponse queryResponseNotInclude = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
                bool isExcludeCellChanges = queryResponseNotInclude.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData) == null;

                Site.Log.Add(
                LogEntryKind.Debug,
                "When Include Cell Changes is not set, the cell manifest is not included. Actually it {0}",
                isExcludeCellChanges ? "is not included" : "is included");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4117
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsTrue(
                             isExcludeCellChanges,
                             "MS-FSSHTTPB",
                             4117,
                             @"[In Appendix B: Product Behavior] If E - Include Cell Changes is not set, cell changes are not included. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
                }
                else
                {
                    this.Site.Assert.IsTrue(isExcludeCellChanges, "[In Query Changes] G - Include Cell Changes (1 bit): otherwise[If E - Include Cell Changes is not set], cell changes are not included");
                }
            }
        }

        /// <summary>
        /// The method uses to verify whether no scoping restriction is specified when the Cell ID is 0x0000.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC21_QueryChanges_CellID()
        {
            // Initialize the service on OneNote file.
            string fileUrl = Common.GetConfigurationPropertyValue("OneNoteFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server with Cell Id set to specified value.
            FsshttpbCellRequest cellRequestFirst = SharedTestSuiteHelper.CreateFsshttpbCellRequest();

            // Initialize a cell id defined in the MS-FSSHTTPD
            CellID cellId = new CellID(new ExGuid(1, new Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073")), new ExGuid(1, new Guid("6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B")));
            QueryChangesCellSubRequest queryChangeScope = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, cellId, null, null, null);
            cellRequestFirst.AddSubRequest(queryChangeScope, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestFirst.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first QueryChanges operation should succeed.");
            FsshttpbResponse queryResponseInScope = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponseInScope, this.Site);

            // Query changes from the protocol server with Cell Id set to 0x0000.
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChangeNoScope = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, false, 0, true, true, 0, null, null, null, null);
            cellRequestSecond.AddSubRequest(queryChangeNoScope, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second QueryChanges operation should succeed.");
            FsshttpbResponse queryResponseNoScope = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            bool isVerifiedR927 = queryResponseInScope.DataElementPackage.DataElements.Count < queryResponseNoScope.DataElementPackage.DataElements.Count;

            Site.Log.Add(
                LogEntryKind.Debug,
                "When query OneNote file with the Cell ID constrained equals 0x0000, the server responds data element count {0}",
                queryResponseNoScope.DataElementPackage.DataElements.Count);

            Site.Log.Add(
                LogEntryKind.Debug,
                "When query OneNote file with the Cell ID constrained does not equal 0x0000, the server responds data element count {0}",
                queryResponseInScope.DataElementPackage.DataElements.Count);

            Site.Log.Add(
                LogEntryKind.Debug,
                "Expect the data elements count is larger when query OneNote file with the Cell ID constrained does not equal 0x0000. Actually it {0}",
                isVerifiedR927 ? "does" : "does not");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R927
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR927,
                         "MS-FSSHTTPB",
                         927,
                         @"[In Query Changes] Cell ID (variable): If the Cell ID is 0x0000, no scoping restriction is specified.");
            }
            else
            {
                this.Site.Assert.IsTrue(isVerifiedR927, "[In Query Changes] Cell ID (variable): If the Cell ID is 0x0000, no scoping restriction is specified.");
            }
        }

        /// <summary>
        /// The method uses to verify server must return same response whenever the C - Exclude Object Data field is set to 0 or 1.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC26_QueryChanges_ExcludeObjectData()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes with Exclude Object Data setting to value 1.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.ExcludeObjectData = 1;
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            // Query changes with Exclude Object Data setting to value 0.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.ExcludeObjectData = 0;
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType querySubResponse2 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse2.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse2 = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse2, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse2, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R214601
                Site.CaptureRequirementIfAreEqual<int>(
                    fsshttpbResponse.DataElementPackage.DataElements.Count,
                    fsshttpbResponse2.DataElementPackage.DataElements.Count,
                    "MS-FSSHTTPB",
                    214601,
                    @"[In Query Changes] Whenever the C –Exclude Object Data field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.AreEqual<int>(
                    fsshttpbResponse.DataElementPackage.DataElements.Count,
                    fsshttpbResponse2.DataElementPackage.DataElements.Count,
                    "Server must return same response whenever the C- Exclude Object Data field is set to 0 or 1.");
            }
        }

        /// <summary>
        /// This method is used to test query changes with the allow fragment 2 flag is false.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC27_QueryChanges_AllowFragments2_Zero()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1348, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Allow Fragments 2 flag.");
            }

            // Initialize the service
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments B flag with the value false.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, null, 10000, null, null);
            // Create query changes request set allow fragments2 E flag with the value false.
            queryChange.AllowFragments2 = 0;
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            DataElement fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsNull(
                    fragDataElement,
                    "MS-FSSHTTPB",
                    1348,
                    @"[In Appendix B: Product Behavior]If E ?Allow Fragments 2 is not set, the Storage Manifest does not allow fragments, unless the bit specified in B is set. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }
            else
            {
                this.Site.Assert.IsNull(
                    fragDataElement,
                    @"[In Appendix B: Product Behavior]If E ?Allow Fragments 2 is not set, the Storage Manifest does not allow fragments, unless the bit specified in B is set. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }

            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            // Create query changes request with allow fragments B flag with the value true.
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, 10000, null, null);
            // Create query changes request set allow fragments2 E flag with the value false.
            queryChange.AllowFragments2 = 0;
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsNotNull(
                    fragDataElement,
                    "MS-FSSHTTPB",
                    1348,
                    @"[In Appendix B: Product Behavior]If E ?Allow Fragments 2 is not set, the Storage Manifest does not allow fragments, unless the bit specified in B is set. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }
            else
            {
                this.Site.Assert.IsNotNull(
                    fragDataElement,
                    @"[In Appendix B: Product Behavior]If E ?Allow Fragments 2 is not set, the Storage Manifest does not allow fragments, unless the bit specified in B is set. (Microsoft Office 2013 and Microsoft SharePoint 2013 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This method is used to test query changes with the allow fragment 2 flag is true.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC28_QueryChanges_AllowFragments2_One()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1348, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support Allow Fragments 2 flag.");
            }

            // Initialize the service
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments B flag with the value false.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, null, 10000, null, null);
            // Create query changes request set allow fragments2 E flag with the value true.
            queryChange.AllowFragments2 = 1;
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");

            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);

            DataElement fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);
            
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsNotNull(
                    fragDataElement,
                    "MS-FSSHTTPB",
                    4040,
                    @"[In Query Changes] E ?Allow Fragments 2 (1 bit): If set, a bit that specifies to allow fragments;( Microsoft SharePoint Server 2013 and above follow this behavior.)");
            }
            else
            {
                this.Site.Assert.IsNotNull(
                    fragDataElement,
                    @"[In Query Changes] E ?Allow Fragments 2 (1 bit): If set, a bit that specifies to allow fragments;( Microsoft SharePoint Server 2013 and above follow this behavior.)");
            }

            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            // Create query changes request with allow fragments B flag with the value true.
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, 10000, null, null);
            // Create query changes request set allow fragments2 E flag with the value true.
            queryChange.AllowFragments2 = 1;
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
            fragDataElement = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.FragmentDataElementData);
            
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsNotNull(
                    fragDataElement,
                    "MS-FSSHTTPB",
                    4040,
                    @"[In Query Changes] E ?Allow Fragments 2 (1 bit): If set, a bit that specifies to allow fragments;( Microsoft SharePoint Server 2013 and above follow this behavior.)");
            }
            else
            {
                this.Site.Assert.IsNotNull(
                    fragDataElement,
                    @"[In Query Changes] E ?Allow Fragments 2 (1 bit): If set, a bit that specifies to allow fragments;( Microsoft SharePoint Server 2013 and above follow this behavior.)");
            }
        }
        #endregion

        #region Knowledge Related
        /// <summary>
        /// This test method aims to verify GUID Combined with the From sequence number forms the starting serial number of the range.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC22_QueryChanges_StartingSerialNumber()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments flag with the value true.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            CellKnowledge cellSpecializedKnowledgeData = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges[0].SpecializedKnowledgeData as CellKnowledge;

            bool isCombined = SharedTestSuiteHelper.CheckFromSequenceNumber(fsshttpbResponse, cellSpecializedKnowledgeData);
            Site.Log.Add(
                LogEntryKind.Debug,
                "Combined with the From sequence number, it [GUID (16 bytes)] forms the starting serial number. Actually it {0}",
                isCombined ? "does" : "does not");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2126
                Site.CaptureRequirementIfIsTrue(
                         isCombined,
                         "MS-FSSHTTPB",
                         2126,
                         @"[In Cell Knowledge Range] GUID (16 bytes): Combined with the From sequence number, it[GUID (16 bytes)] forms the starting Serial Number (section 2.2.1.9) of the range.");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2128
                Site.CaptureRequirementIfIsTrue(
                         isCombined,
                         "MS-FSSHTTPB",
                         2128,
                         @"[In Cell Knowledge Range] From (variable): When combined with the GUID, it[From (variable)] forms the serial number of the starting data element in the range.");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4007
                Site.CaptureRequirementIfIsNotNull(
                         cellSpecializedKnowledgeData.CellKnowledgeRangeList,
                         "MS-FSSHTTPB",
                         4007,
                         @"[In Serial Number] The server will return a Cell Knowledge Range that specifies the range of serial numbers, as specified in section 2.2.1.13.2.1.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isCombined,
                    @"[In Cell Knowledge Range] GUID (16 bytes): Combined with the From sequence number, it[GUID (16 bytes)] forms the starting serial number (section 2.2.1.9) of the range.");

                Site.Assert.IsNotNull(
                    cellSpecializedKnowledgeData.CellKnowledgeRangeList,
                    @"[In Serial Number] The server will return a Cell Knowledge Range that specifies the range of serial numbers, as specified in section 2.2.1.13.2.1.");
            }
        }

        /// <summary>
        /// This test method aims to verify GUID Combined with the To sequence number forms the ending serial number of the range.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC23_QueryChanges_EndingSerialNumber()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with allow fragments flag with the value true.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            CellKnowledge cellSpecializedKnowledgeData = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge.SpecializedKnowledges[0].SpecializedKnowledgeData as CellKnowledge;
            bool isCombined = SharedTestSuiteHelper.CheckToSequenceNumber(fsshttpbResponse, cellSpecializedKnowledgeData);
            Site.Log.Add(
                LogEntryKind.Debug,
                "Combined with the To sequence number, it[GUID (16 bytes)] forms the ending serial number of the range. Actually it {0}",
                isCombined ? "does" : "does not");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2127
                Site.CaptureRequirementIfIsTrue(
                         isCombined,
                         "MS-FSSHTTPB",
                         2127,
                         @"[In Cell Knowledge Range] GUID (16 bytes): Combined with the To sequence number, it[GUID (16 bytes)] forms the ending Serial Number of the range.");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2129
                Site.CaptureRequirementIfIsTrue(
                         isCombined,
                         "MS-FSSHTTPB",
                         2129,
                         @"[In Cell Knowledge Range] To (variable): When combined with the GUID, it[To (variable)] forms the Serial Number of the ending data element in the range.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isCombined,
                    @"[In Cell Knowledge Range] GUID (16 bytes): Combined with the To sequence number, it[GUID (16 bytes)] forms the ending serial number of the range.");
            }
        }

        /// <summary>
        /// This test method aims to verify the serial number in Waterline Knowledge is greater than or equal to the serial number of all the cells on the server that the client has downloaded.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC24_QueryChanges_Waterline_SerialNumber()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a query change subRequest.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            QueryChangesSubResponseData queryChangeSubResponseData = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>();
            SpecializedKnowledge specializedKnowledge;
            bool isContainedWaterline = (specializedKnowledge = queryChangeSubResponseData.Knowledge.SpecializedKnowledges.FirstOrDefault(specialized => SpecializedKnowledge.WaterlineKnowledgeGuid.Equals(specialized.GUID))) != null;
            this.Site.Assert.IsTrue(isContainedWaterline, "The waterlineKnowledge should be returned.");
            WaterlineKnowledge waterlineKnowledge = specializedKnowledge.SpecializedKnowledgeData as WaterlineKnowledge;
            bool isVerifiedR570 = true;

            var waterlineKnowledgeEntry = waterlineKnowledge.WaterlineKnowledgeData.FirstOrDefault(waterLine => waterLine.Waterline.DecodedValue > 1);
            Site.Assert.IsNotNull(
                waterlineKnowledgeEntry,
                "There should be at least one water line entry with value larger than 1");

            // The storage index data element needs to be filter out.
            foreach (
                var dataElement in
                    fsshttpbResponse.DataElementPackage.DataElements.Where(
                        element => element.DataElementType != DataElementType.StorageIndexDataElementData))
            {
                if (waterlineKnowledgeEntry.Waterline.DecodedValue < dataElement.SerialNumber.Value)
                {
                    isVerifiedR570 = false;

                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "If the water line value is {1}, expect the serialNumer value is less than it, but the serialNumer value is {0}.",
                        dataElement.SerialNumber.Value,
                        waterlineKnowledgeEntry.Waterline.DecodedValue);
                    break;
                }
            }

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R570
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR570,
                         "MS-FSSHTTPB",
                         570,
                         @"[In Waterline Knowledge] The Waterline Knowledge specifies the current server waterline, which is the Serial Number (section 2.2.1.9) greater than or equal to the Serial Number of all the cells on the server that the client has downloaded.");
            }
            else
            {
                this.Site.Assert.IsTrue(isVerifiedR570, "[In Waterline Knowledge] The Waterline Knowledge specifies the current server waterline, which is the serial number (section 2.2.1.9) greater than or equal to the serial number of all the cells on the server that the client has downloaded.");
            }
        }
        #endregion

        /// <summary>
        /// This test method aims to verify server must return the same response when Reserved field is set to 0 or 1.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC25_QueryChanges_ReservedIsIgnored()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create query changes request with setting Reserved to value 0.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChangeWithReserved0 = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, null, null);
            queryChangeWithReserved0.Reserved = 0;
            cellRequest.AddSubRequest(queryChangeWithReserved0, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            // Create query changes request with setting Reserved to value 1.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChangeWithReserved1 = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, null, null);
            queryChangeWithReserved1.Reserved = 1;
            cellRequest.AddSubRequest(queryChangeWithReserved1, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            FsshttpbResponse queryResponse2 = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R43001
                Site.CaptureRequirementIfAreEqual<int>(
                    queryResponse.DataElementPackage.DataElements.Count,
                    queryResponse2.DataElementPackage.DataElements.Count,
                    "MS-FSSHTTPB",
                    43001,
                    @"[In Query Changes] Whenever the A ?Reserved field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.AreEqual<int>(
                    queryResponse.DataElementPackage.DataElements.Count,
                    queryResponse2.DataElementPackage.DataElements.Count,
                    "Server must return same response whenever the A- Reserved field is set to 0 or 1.");
            }
        }


        /// <summary>
        /// This test method aims to verify flag Round Knowledge to Whole Cell Changes.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S12_TC29_QueryChanges_RoundKnowledgeToWholeCellChanges()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query change
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, null, null);
            cellRequest.AddSubRequest(queryChange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            ExGuid storageIndex = queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Put Change
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            dataElements.AddRange(queryResponse.DataElementPackage.DataElements);
            cellRequest.AddSubRequest(putChange, dataElements);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");

            // Query change again
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, null, null);
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            Knowledge knowledge = queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().Knowledge;

            // Query change with knowledge returned in previous step
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, true, false, true, 0, true, true, 0, null, null, null, null);
            queryChange.RoundKnowledgeToWholeCellChanges = 1;
            queryChange.Knowledge = knowledge;
            cellRequest.AddSubRequest(queryChange, null);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the query changes succeed.");
            queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
            DataElement data = queryResponse.DataElementPackage.DataElements.FirstOrDefault(dataElement => dataElement.DataElementType == DataElementType.CellManifestDataElementData);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured && Common.IsRequirementEnabled(1351,this.Site))
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R1351
                Site.CaptureRequirementIfIsNull(
                    data,
                    "MS-FSSHTTPB",
                    1351,
                    @"[In Appendix B: Product Behavior]If set, a bit that specifies that the knowledge specified in the request SHOULD be modified, prior to change enumeration, such that any changes under a cell node, as implied by the knowledge, cause the knowledge to be modified such that all changes in that cell are returned. (Microsoft Office 2016/Microsoft SharePoint 2016 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.IsNull(
                    data,
                    "There should no changes returned if set Knowledge to that has queried and set RoundKnowledgeToWholeCellChanges.");
            }
        }
        #endregion
    }
}