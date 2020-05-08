namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with Properties operation.
    /// </summary>
    [TestClass]
    public abstract class S20_Properties : SharedTestSuiteBase
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
        /// A method used to clean up this class.
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
        public void S20_PropertiesInitialization()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 246801, this.Site), "This test case only runs when Properties subrequest is supported.");
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "Properties" sub-request.

        /// <summary>
        /// A method used to verify that Properties sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S20_TC01_Properties_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string[] ids = new string[] { "RealtimeTypingEndpointUrl", "DocumentAccessToken", "DocumentAccessTokenTtl" };
            PropertyIdType propertyId = new PropertyIdType();
            PropertyIdType[] propertyIds = new PropertyIdType[3];

            for(int i=0; i<ids.Length; i++)
            {
               propertyId.id= ids[i];
               propertyIds[i] = propertyId;
            }

            PropertiesSubRequestType properties = SharedTestSuiteHelper.CreatePropertiesSubRequest(SequenceNumberGenerator.GetCurrentToken(), PropertiesRequestTypes.PropertyGet, propertyIds, this.Site);
            
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { properties });
            PropertiesSubResponseType propertiesResponse = SharedTestSuiteHelper.ExtractSubResponse<PropertiesSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            SubResponseType subresponse = cellStorageResponse.ResponseCollection.Response[0].SubResponse[0];

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture the requirement MS-FSSHTTP_R246801
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         246801,
                         @"[In Appendix B: Product Behavior] Implementation does support Properties operation. (SharePoint Server 2016 and above follow this behavior.)");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2301011
                Site.CaptureRequirementIfAreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "MS-FSSHTTP",
                    2301011,
                    @"[PropertiesSubResponseType]In the case of success, it contains information requested as part of a Properties subrequest. ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2443
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2443,
                         @"[Properties Subrequest][The protocol server returns results based on the following conditions:]An ErrorCode value of ""Success"" indicates success in processing the Properties request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2447
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         2447,
                         @"[Property Get]If the Properties attribute is set to ""PropertyGet"", the protocol server considers the Properties subrequest to be of type ""Property Get "". ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2299
                Site.CaptureRequirementIfIsNotNull(
                    propertiesResponse.SubResponseData.PropertyValues,
                    "MS-FSSHTTP",
                    2299,
                    @"[PropertiesSubResponseDataType][PropertyValues]This element MUST only be included in the response if the Properties attribute value is set to ""PropertyGet"".");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "[PropertiesSubResponseType]In the case of success, it contains information requested as part of a Properties subrequest. ");
            }
        }
        #endregion
    }
}