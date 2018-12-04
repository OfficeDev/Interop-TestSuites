namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the schema lock sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S03_SchemaLock : S03_SchemaLock
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S03_SchemaLock.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S03_SchemaLock.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify the related requirements when get a schema lock on a file which is not existent.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S03_TC01_GetLock_FileNotExistsOrCannotBeCreated()
        {
            // Get a schema lock with a nonexistent file, expect the server returns the error code "FileNotExistsOrCannotBeCreated".
            string fileUrlNotExit = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the service
            this.InitializeContext(fileUrlNotExit, this.UserName01, this.Password01, this.Domain);

            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(fileUrlNotExit, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1913
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotExistsOrCannotBeCreated,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1913,
                         @"[In SchemaLock Subrequest][The protocol server returns results based on the following conditions:] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotExistsOrCannotBeCreated,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SchemaLock Subrequest][The protocol server returns results based on the following conditions:] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S03_TC02_GetLock_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a GetLock subRequest with all valid parameters.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3006Verified = false;
            try
            {
                // Send a GetLock for schema lock subRequest to the protocol server without specifying URL attribute.
                response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3006Verified = message.Contains("Duplicate attribute");
                isR3006Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3006
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3006Verified,
                             "MS-FSSHTTP",
                             3006,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3007
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3007,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.Assert.IsTrue(
                        isR3006Verified,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S03_TC03_GetLock_UrlEmpty()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a GetLock subRequest with all valid parameters.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3008Verified = false;
            try
            {
                // Send a GetLock for schema lock subRequest to the protocol server without specifying URL attribute.
                response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3008Verified = message.Contains("Duplicate attribute");
                isR3008Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3008
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty string.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3008Verified,
                             "MS-FSSHTTP",
                             3008,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty string.");

                    Site.Assert.IsTrue(
                        isR3008Verified,
                        "[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. &lt;3&gt; Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when getting a schema Lock on a file which is already checked out.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S03_TC04_GetLock_Success_ExclusiveLockReturnReason_CheckedOutByCurrentUser()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out one file by a specified user name.
            bool isCheckOutSuccess = SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            Site.Assert.AreEqual(true, isCheckOutSuccess, "Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Change the authentication mode
                if (!this.SutPowerShellAdapter.SwitchClaimsAuthentication(false))
                {
                    this.Site.Assert.Fail("Cannot change the authentication mode to windows based.");
                }

                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }

            CheckLockAvailability();

            // Get a schema lock with AllowFallbackToExclusive set to true, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, true, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "The test case cannot continue unless the server responses Success when the user get a schema lock with AllowFallBack attribute true on an already checked out file by the same account.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            #region Verify related requirements

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R404
                Site.CaptureRequirementIfAreEqual<LockTypes>(
                         LockTypes.ExclusiveLock,
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         404,
                         @"[In LockTypes] ExclusiveLock: The string value ""ExclusiveLock"", indicating an exclusive lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R406
                Site.CaptureRequirementIfAreEqual<LockTypes>(
                         LockTypes.ExclusiveLock,
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         406,
                         @"[In LockTypes][ExclusiveLock] In a cell storage service response message, an exclusive lock indicates that an exclusive lock is granted to the current client for that specific file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1428
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R1428 and MS-FSSHTTP_R717, the ExclusiveLockReturnReason attribute in the response should be specified, actually it is {0}.",
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified ? "specified" : "not specified");

                Site.CaptureRequirementIfIsTrue(
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         1428,
                         @"[In SubResponseDataOptionalAttributes][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations when the LockType attribute in the subresponse is set to ""ExclusiveLock"": A schema lock subrequest of type ""Get lock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R717
                Site.CaptureRequirementIfIsTrue(
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         717,
                         @"[In SchemaLockSubResponseDataType] The ExclusiveLockReturnReason attribute MUST be specified in a schema lock subresponse that is generated in response to a schema lock subrequest of type ""Get lock"" when the LockType attribute in the subresponse is set to ""ExclusiveLock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R715
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CheckedOutByCurrentUser,
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         715,
                         @"[In SchemaLockSubResponseDataType] ExclusiveLockReturnReason: An ExclusiveLockReturnReasonTypes that specifies the reason why an exclusive lock is granted in a schema lock subresponse.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R350
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CheckedOutByCurrentUser,
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         350,
                         @"[In ExclusiveLockReturnReasonTypes] CheckedOutByCurrentUser: The string value ""CheckedOutByCurrentUser"", indicating that an exclusive lock is granted on the file because the file is checked out by the current user who sent the cell storage service request message.");
            }
            else
            {
                Site.Assert.AreEqual<LockTypes>(
                    LockTypes.ExclusiveLock,
                    schemaLockSubResponse.SubResponseData.LockType,
                    @"[In LockTypes] ExclusiveLock: The string value ""ExclusiveLock"", indicating an exclusive lock on the file.");

                Site.Assert.IsTrue(
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                    @"[In SubResponseDataOptionalAttributes][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations when the LockType attribute in the subresponse is set to ""ExclusiveLock"": A schema lock subrequest of type ""Get lock"".");

                Site.Assert.AreEqual<ExclusiveLockReturnReasonTypes>(
                    ExclusiveLockReturnReasonTypes.CheckedOutByCurrentUser,
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReason,
                    @"[In SchemaLockSubResponseDataType] ExclusiveLockReturnReason: An ExclusiveLockReturnReasonTypes that specifies the reason why an exclusive lock is granted in a schema lock subresponse.");
            }

            #endregion
        }

        /// <summary>
        /// Initialize the shared context based on the specified request file URL, user name, password and domain for the MS-FSSHTTP test purpose.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        protected override void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            SharedContextUtils.InitializeSharedContextForFSSHTTP(userName, password, domain, this.Site);
        }

        /// <summary>
        /// Merge the common configuration and should/may configuration file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        protected override void MergeConfigurationFile(TestTools.ITestSite site)
        {
            ConfigurationFileHelper.MergeConfigurationFile(site);
        }
    }
}