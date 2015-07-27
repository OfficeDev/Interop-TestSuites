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
    /// A class which contains test cases used to capture the requirements related with subResponse to multiple subRequests.
    /// </summary>
    [TestClass]
    public abstract class S10_MultipleSubRequests : SharedTestSuiteBase
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
        public void S10_MultipleSubRequestsInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should be unique for each test case
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Case

        /// <summary>
        /// A method used to test the subRequest token mapping between the subRequest and subResponse.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC01_SubRequestToken()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Update the file contents when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            CellSubResponseType putChangeResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request succeeds.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, exclusiveLocksubRequest.SubRequestData.ExclusiveLockID);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putChangeResponse.ErrorCode, this.Site), "Test case cannot continue unless the put changes sub request succeeds.");

            this.Site.Assert.AreEqual<int>(
                                int.Parse(exclusiveLocksubRequest.SubRequestToken),
                                int.Parse(exclusiveResponse.SubRequestToken),
                                "Test case cannot run unless the put change response token equals the expected sub request token.");
            this.Site.Assert.AreEqual<int>(
                               int.Parse(putChange.SubRequestToken),
                               int.Parse(putChangeResponse.SubRequestToken),
                               "Test case cannot run unless the put change response token equals the expected sub request token.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If both the subRequest tokens are mapping, then MS-FSSHTTP_R1484 and MS-FSSHTTP_R283 can be captured.
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1484,
                         @"[In SubResponseType] SubRequestToken: If client sends 2 SubRequest elements in the Request, in the server response, SubRequestToken uniquely identifies the 2 SubRequest element whose subresponse is being generated as part of the SubResponse element.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R283
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         283,
                         @"[In SubResponseType][SubRequestToken] The mapping subresponse that gets generated for the subrequest references the SubRequestToken to indicate that it is the response for that subrequest.");
            }
        }

        /// <summary>
        /// A method used to test two subRequests procession when the dependency type is OnSuccess.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC02_DependencyTypeOnSuccess()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Update the file contents when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnSuccess;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            CellSubResponseType putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request succeeds.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, exclusiveLocksubRequest.SubRequestData.ExclusiveLockID);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site), "Test case cannot continue unless the put changes sub request succeeds.");

            // Make another GetLock request with a different exclusive lock ID.
            exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            exclusiveLocksubRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Also make another putChange request dependent on the new GetLock request
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnSuccess;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request fails.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the first putChanges request succeed and the second putChanges request fails, then capture requirement R333.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         333,
                         @"[In DependencyTypes] OnSuccess: Indicates that the subrequest MUST be processed only on the successful execution of the other subrequest.");

                // If the error code equals DependentOnlyOnSuccessRequestFailed, then capture R322
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         322,
                         @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnSuccessRequestFailed: Indicates an error when the subrequest on which this specific subrequest is dependent has failed and the DependencyType attribute in this subrequest is set to ""OnSuccess"" or [""OnSuccessOrNotSupported""].");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnSuccess: Indicates that the subrequest MUST be processed only on the successful execution of the other subrequest.");
            }
        }

        /// <summary>
        /// A method used to test two subRequests processing when the dependency type is OnFail.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC03_DependencyTypeOnFail()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Update the file contents when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnFail;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request succeeds.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, exclusiveLocksubRequest.SubRequestData.ExclusiveLockID);

            CellSubResponseType putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R323
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnFailRequestSucceeded,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         323,
                         @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnFailRequestSucceeded: 
                         Indicates an error when the subrequest on which this specific subrequest is dependent has succeeded and the DependencyType attribute in this subrequest is set to ""OnFail"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DependentOnlyOnFailRequestSucceeded,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                    @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnFailRequestSucceeded: 
                        Indicates an error when the subrequest on which this specific subrequest is dependent has succeeded and the DependencyType attribute in this subrequest is set to ""OnFail"".");
            }

            // Make another GetLock request with a different exclusive lock ID.
            exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            exclusiveLocksubRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Also make another putChange request dependent on the new GetLock request.
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnFail;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request fails.");
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site), "Test case cannot continue unless the put changes sub request succeeds.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R334
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         334,
                         @"[In DependencyTypes] OnFail: Indicates that the subrequest MUST be processed only on the failed execution of the other subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnFail: Indicates that the subrequest MUST be processed only on the failed execution of the other subrequest.");
            }
        }

        /// <summary>
        /// A method used to test two sub requests and the dependency type is OnNotSupported.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC04_DependencyTypeOnNotSupported_Support()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Update the file contents when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is OnNotSupported.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnNotSupported;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            CellSubResponseType putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request succeeds.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, exclusiveLocksubRequest.SubRequestData.ExclusiveLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals DependentOnlyOnNotSupportedRequestGetSupported, then capture R324 and R2244
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnNotSupportedRequestGetSupported,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         324,
                         @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnNotSupportedRequestGetSupported:
                         Indicates an error when the subrequest on which this specific subrequest is dependent is supported and the DependencyType attribute in this subrequest is set to ""OnNotSupported"" or [""OnSuccessOrOnNotSupported""].");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2244
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnNotSupportedRequestGetSupported,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2244,
                         @"[In DependencyTypes] OnNotSupported: Indicates that the subrequest MUST NOT be processed if the other subrequest is supported.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DependentOnlyOnNotSupportedRequestGetSupported,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                    @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnNotSupportedRequestGetSupported:
                        Indicates an error when the subrequest on which this specific subrequest is dependent is supported and the DependencyType attribute in this subrequest is set to ""OnNotSupported"" or [""OnSuccessOrOnNotSupported""].");
            }
        }

        /// <summary>
        /// A method used to test two subRequests when the dependency type is OnSuccessOrNotSupported.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC05_DependencyTypeOnSuccessOrNotSupported_OnSuccess()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Update the file contents when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnSuccessOrNotSupported;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            CellSubResponseType putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request succeeds.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, exclusiveLocksubRequest.SubRequestData.ExclusiveLockID);

            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site), "Test case cannot continue unless the put changes sub request succeeds.");

            // Make another GetLock request with a different exclusive lock ID.
            exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            exclusiveLocksubRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Also make another putChange request dependent on the new GetLock request.
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10)));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            // Make a putChanges subRequest dependent on the exclusive lock and the dependency type is success.
            putChange.DependencyTypeSpecified = true;
            putChange.DependencyType = DependencyTypes.OnSuccessOrNotSupported;
            putChange.DependsOn = exclusiveLocksubRequest.SubRequestToken;

            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest, putChange });

            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            putChanges = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 1, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "Test case cannot continue unless the get lock sub request fails.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the first putChanges request succeed and the second putChanges request fails, then capture requirement R336 and R32201.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         336,
                         @"[In DependencyTypes] OnSuccessOrNotSupported: Indicates that the subrequest MUST be processed only when one of the following conditions is true:
                         On the successful execution of the other subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R32201
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         32201,
                         @"[In DependencyCheckRelatedErrorCodeTypes] DependentOnlyOnSuccessRequestFailed: Indicates an error when the subrequest on which this specific subrequest is dependent has failed and the DependencyType attribute in this subrequest is set to [""OnSuccess""] or ""OnSuccessOrNotSupported"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DependentOnlyOnSuccessRequestFailed,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(putChanges.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnSuccessOrNotSupported: Indicates that the subrequest MUST be processed only when one of the following conditions is true:
                    On the successful execution of the other subrequest.");
            }
        }

        /// <summary>
        /// A method used to test two subRequests when the dependency type is OnNotSupported and the dependent subRequest is not supported.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC06_DependencyTypeOnNotSupported_NotSupport()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 2243, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not return Success when its sub-request dependency type is OnNotSupported and the dependent sub-request is not supported.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the coauthoring feature so as to create a unsupported coauthoring subRequest.
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Create a coauthoring subRequest with all valid parameters.
            CoauthSubRequestType coauthSubRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Create a ServerTime subRequest which depends on the coauthoring subRequest.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            serverTimeSubRequest.DependencyType = DependencyTypes.OnNotSupported;
            serverTimeSubRequest.DependencyTypeSpecified = true;
            serverTimeSubRequest.DependsOn = coauthSubRequest.SubRequestToken;

            // Send the two subRequests to the protocol server, expect the coauthoring operation fails.
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { coauthSubRequest, serverTimeSubRequest });
            CoauthSubResponseType coauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(coauthSubResponse.ErrorCode, this.Site), "The coauthoring operation should fail.");

            ServerTimeSubResponseType serverTimeSubResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(response, 0, 1, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2243
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2243,
                         @"[In DependencyTypes] OnNotSupported: Indicates that the subrequest MUST be processed if the other subrequest is not supported. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnNotSupported: Indicates that the subrequest MUST be processed if the other subrequest is not supported. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
        }

        /// <summary>
        /// A method used to test the protocol server returns error code "InvalidRequestDependencyType" when a subRequest dependency type that is not valid is specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC07_InvalidRequestDependencyType()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 325, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not return InvalidRequestDependencyType when the sub-request dependency type that is not valid.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a serverTime subRequest with all valid parameters.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());

            // Create a WhoAmI subRequest with all valid parameters and depends on the serverTime subRequest.
            WhoAmISubRequestType whoAmiSubRequest = SharedTestSuiteHelper.CreateWhoAmISubRequest(SequenceNumberGenerator.GetCurrentToken());
            whoAmiSubRequest.DependencyType = DependencyTypes.Invalid;
            whoAmiSubRequest.DependencyTypeSpecified = true;
            whoAmiSubRequest.DependsOn = serverTimeSubRequest.SubRequestToken;

            // Send the subRequest to the protocol server, expect the protocol server returns error code "InvalidRequestDependencyType".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { serverTimeSubRequest, whoAmiSubRequest });
            ServerTimeSubResponseType serverTimeSubResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site), "The serverTime operation should succeed.");
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 1, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R325
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidRequestDependencyType,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         325,
                         @"[In DependencyCheckRelatedErrorCodeTypes] InvalidRequestDependencyType: 
                             Indicates an error when a subrequest dependency type that is not valid is specified. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidRequestDependencyType,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In DependencyCheckRelatedErrorCodeTypes] InvalidRequestDependencyType: 
                        Indicates an error when a subrequest dependency type that is not valid is specified. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
        }

        /// <summary>
        /// A method used to test two subRequests when the dependency type is OnSuccessOrNotSupported and the dependent subRequest is not supported.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC08_OnSuccessOrNotSupported_OnNotSupport()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 337, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not return Success when its sub-request dependency type is OnSuccessOrNotSupported and the dependent subRequest is not supported");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the coauthoring feature.
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // create a coauthoring subRequest with all valid parameters.
            CoauthSubRequestType coauthSubRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Create a serverTime subRequest with DependencyType set to "OnSuccessOrNotSupported", which depends on the coauthoring subRequest.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            serverTimeSubRequest.DependencyType = DependencyTypes.OnSuccessOrNotSupported;
            serverTimeSubRequest.DependencyTypeSpecified = true;
            serverTimeSubRequest.DependsOn = coauthSubRequest.SubRequestToken;

            // Send these two subRequests to the protocol server, expect the protocol server returns error code "Success" for the serverTimeSubRequest.
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { coauthSubRequest, serverTimeSubRequest });
            CoauthSubResponseType coauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(coauthSubResponse.ErrorCode, this.Site), "The coauthoring operation should be failed.");
            ServerTimeSubResponseType serverTimeSubResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(response, 0, 1, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R337
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         337,
                         @"[In DependencyTypes] OnSuccessOrNotSupported: Indicates that the subrequest MUST be processed only when one of the following conditions is true:
                             If the other subrequest is not supported. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnSuccessOrNotSupported: Indicates that the subrequest MUST be processed only when one of the following conditions is true:
                    If the other subrequest is not supported. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
            }
        }

        /// <summary>
        /// A method used to test two sub requests when the dependency type is OnExecute and the dependent sub Request is not executed.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S10_TC09_DependencyTypeOnExecute()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a coauthoring subRequest with all valid parameters.
            CoauthSubRequestType firstCoauthSubRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Create another coauthoring subRequest which depends on the previous coauthoring subRequest and set the dependency type to OnFail so as to make the second subRequest not executed.
            CoauthSubRequestType secondCoauthSubRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID);
            secondCoauthSubRequest.DependencyType = DependencyTypes.OnFail;
            secondCoauthSubRequest.DependencyTypeSpecified = true;
            secondCoauthSubRequest.DependsOn = firstCoauthSubRequest.SubRequestToken;

            // Create a serverTime subRequest which depends on the second coauthoring subRequest and set the dependency type to OnExecute.
            // Create a ServerTime subRequest which depends on the coauthoring subRequest.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            serverTimeSubRequest.DependsOn = secondCoauthSubRequest.SubRequestToken;
            serverTimeSubRequest.DependencyType = DependencyTypes.OnExecute;
            serverTimeSubRequest.DependencyTypeSpecified = true;

            // Send these three subRequests to the protocol server, expect the protocol server returns error code "Success" for the first coauthoring subRequest, returns "DependentOnlyOnFailRequestSucceeded" for the second coauthoring subRequest and returns "DependentRequestNotExecuted" for the severTime subRequest.
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { firstCoauthSubRequest, secondCoauthSubRequest, serverTimeSubRequest });
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CoauthSubResponseType firstCoauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(firstCoauthSubResponse.ErrorCode, this.Site), "The first coauthoring operation should succeed.");
            CoauthSubResponseType secondCoauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 1, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.DependentOnlyOnFailRequestSucceeded, SharedTestSuiteHelper.ConvertToErrorCodeType(secondCoauthSubResponse.ErrorCode, this.Site), "The second coauthoring operation should not be executed.");
            ServerTimeSubResponseType serverTimeSubResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(response, 0, 2, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 332, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R332
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.DependentRequestNotExecuted,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             332,
                             @"[In DependencyTypes] OnExecute:Indicates that the subrequest MUST be processed only on the execution of the other subrequest. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 follow this behavior.)");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.DependentRequestNotExecuted,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             321,
                             @"[In DependencyCheckRelatedErrorCodeTypes] DependentRequestNotExecuted: 
                             Indicates an error when the subrequest on which this specific subrequest is dependent has not been executed and the DependencyType attribute in this subrequest is set to ""OnExecute"".");
                }
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DependentRequestNotExecuted,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                    @"[In DependencyTypes] OnExecute:Indicates that the subrequest MUST be processed only on the execution of the other subrequest. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 follow this behavior.)");
            }
        }
        #endregion
    }
}