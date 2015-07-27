//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with AllocateExtendedGuidRange operation.
    /// </summary>
    [TestClass]
    public abstract class S14_AllocateExtendedGuidRange : SharedTestSuiteBase
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// A method used to initialize this class.
        /// </summary>
        /// <param name="testContext">A parameter represents the context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            SharedTestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// A method used to clean up the test environment.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            SharedTestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test Case Initialization

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S14_AllocateExtendedGuidRangeInitialization()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99099, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the AllocateExtendedGuidRange operation.");
            }

            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// A method used to test the Allocate ExtendedGuid Range subRequest processing.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S14_TC01_AllocateExtendedGuidRange_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Allocate Extended GUID Range
            Compact64bitInt requestIdCount = new Compact64bitInt(8000);
            CellSubRequestType allocateRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedAllocateExtendedGuidRange((int)SequenceNumberGenerator.GetCurrentToken(), SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), requestIdCount);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { allocateRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the allocate extended guid range operation on the file {0} succeed.",
                    this.DefaultFileUrl);

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            // Extract the Allocate ExtendedGuid Range sub-response info
            this.Site.Assert.IsTrue(fsshttpbResponse.CellSubResponses.Count > 0, "The protocol server should return SubResponse for Allocate ExtendedGuid Range sub-request processing.");
            this.Site.Assert.IsNotNull(fsshttpbResponse.CellSubResponses[0].GetSubResponseData<AllocateExtendedGuidRangeSubResponseData>(), "The protocol server should return SubResponseData for Allocate ExtendedGuid Range sub-request processing.");
            
            AllocateExtendedGuidRangeSubResponseData allocateResponse = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<AllocateExtendedGuidRangeSubResponseData>();

            // Assert for elements in Allocate Extended Guid Range sub-response
            this.Site.Assert.IsNotNull(allocateResponse.AllocateExtendedGUIDRangeResponse, "The protocol server should return Allocate Extended GUID Range Response.");
            this.Site.Assert.IsNotNull(allocateResponse.GUIDComponent, "The protocol server should allocate ExtendedGuids as requested.");
            this.Site.Assert.IsNotNull(allocateResponse.IntegerRangeMin, "The protocol server should return an integer Range Min in the sub-response.");
            this.Site.Assert.IsNotNull(allocateResponse.IntegerRangeMax, "The protocol server should return an integer Range Max in the sub-response.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R9005
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         99099,
                         @"[In Appendix B: Product Behavior] The implementation does support this sub-request [Allocate ExtendedGuid Range Sub-Request]( SharePoint Server 2013 and above follow this behavior.)");
            }

            // Allocate Extended GUID Range value less than 1000.
            requestIdCount = new Compact64bitInt(100);
            allocateRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedAllocateExtendedGuidRange((int)SequenceNumberGenerator.GetCurrentToken(), SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), requestIdCount);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { allocateRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the allocate extended guid range operation on the file {0} succeed.",
                    this.DefaultFileUrl);

            // Allocate Extended GUID Range value less than 100000.
            requestIdCount = new Compact64bitInt(150000);
            allocateRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedAllocateExtendedGuidRange((int)SequenceNumberGenerator.GetCurrentToken(), SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), requestIdCount);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { allocateRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the allocate extended guid range operation on the file {0} succeed.",
                    this.DefaultFileUrl);
        }

        /// <summary>
        /// A method used to test the protocol server returns rough same allocate extended GUID range response when the reserved value is different in the request.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S14_TC02_AllocateExtendedGuidRange_ReservedIgnore()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            FsshttpbCellRequest fsshttpRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            AllocateExtendedGuidRangeCellSubRequest allocateExtendedGuidRange = new AllocateExtendedGuidRangeCellSubRequest(new Compact64bitInt(8000), SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());

            // Make the reserved value equal to 0.
            allocateExtendedGuidRange.Reserved = 0;
            fsshttpRequest.AddSubRequest(allocateExtendedGuidRange, null);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest((ulong)SequenceNumberGenerator.GetCurrentToken(), fsshttpRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the allocate extended guid range operation on the file {0} succeed.",
                    this.DefaultFileUrl);

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            FsshttpbCellRequest fsshttpRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            AllocateExtendedGuidRangeCellSubRequest allocateExtendedGuidRangeSecond = new AllocateExtendedGuidRangeCellSubRequest(new Compact64bitInt(8000), SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());

            // Make the reserved value equal to 1.
            allocateExtendedGuidRangeSecond.Reserved = 1;
            fsshttpRequestSecond.AddSubRequest(allocateExtendedGuidRangeSecond, null);
            CellSubRequestType cellSubRequestSecond = SharedTestSuiteHelper.CreateCellSubRequest((ulong)SequenceNumberGenerator.GetCurrentToken(), fsshttpRequestSecond.ToBase64());
            CellStorageResponse responseSecond = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequestSecond });
            CellSubResponseType cellSubResponseSecond = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(responseSecond, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseSecond.ErrorCode, this.Site),
                    "Test case cannot continue unless the allocate extended guid range operation on the file {0} succeed.",
                    this.DefaultFileUrl);

            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponseSecond, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponseSecond, this.Site);

            // Compare this two responses roughly, and if these two responses are identical in main part, then capture the requirement MS-FSSHTTPB_R2189
            bool isVerifyR2189 = SharedTestSuiteHelper.ComapreSucceedFsshttpAllocateExtendedGuidRangeResposne(fsshttpbResponse, fsshttpbResponseSecond, this.Site);

            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two allocate extended GUID range responses are same, actual {0}", isVerifyR2189);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2189
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR2189,
                         "MS-FSSHTTPB",
                         2189,
                         @"[In Allocate Extended GUID Range] A - Reserved (8 bits): Whenever A - Reserved is set to one or zero, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isVerifyR2189,
                    @"[In Allocate Extended GUID Range] A - Reserved (8 bits): Whenever A - Reserved is set to one or zero, the protocol server must return the same response.");
            }
        }
        #endregion 
    }
}