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
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with GetDocMetaInfo operation.
    /// </summary>
    [TestClass]
    public abstract class S08_GetDocMetaInfo : SharedTestSuiteBase
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
        public void S08_GetDocMetaInfoInitialization()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9003, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetDocMetaInfo operation.");
            }

            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases for "GetDocMetaInfo" sub-request.

        /// <summary>
        /// A method used to verify that GetDocMetaInfo sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S08_TC01_GetDocMetaInfo_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetDocMetaInfo"sub-request with correct input parameters.
            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = SharedTestSuiteHelper.CreateGetDocMetaInfoSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getDocMetaInfoSubRequest });
            GetDocMetaInfoSubResponseType getDocMetaInfoSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetDocMetaInfoSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getDocMetaInfoSubResponse, "The object 'getDocMetaInfoSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getDocMetaInfoSubResponse.ErrorCode, "The object 'getDocMetaInfoSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the value of "ErrorCode" in the sub-response equals "Success", then capture MS-FSSHTTP_R2016, and MS-FSSHTTP_R1802.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getDocMetaInfoSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2016,
                         @"[In GetDocMetaInfo Subrequest][The protocol returns results based on the following conditions:] Otherwise[the processing of the GetDocMetaInfo subrequest by the protocol server get the requested metadata successfully], the protocol server sets the error code value to ""Success"" to indicate success in processing the GetDocMetaInfo subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1802
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getDocMetaInfoSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1802,
                         @"[In GetDocMetaInfoSubResponseType] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the GetDocMetaInfo subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R9003
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         9003,
                         @"[In Appendix B: Product Behavior] Implementation does support GetDocMetaInfo operation. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getDocMetaInfoSubResponse.ErrorCode, this.Site),
                    @"[In GetDocMetaInfo Subrequest][The protocol returns results based on the following conditions:] Otherwise[the processing of the GetDocMetaInfo subrequest by the protocol server get the requested metadata successfully], the protocol server sets the error code value to ""Success"" to indicate success in processing the GetDocMetaInfo subrequest.");
            }
        }

        #endregion 
    }
}