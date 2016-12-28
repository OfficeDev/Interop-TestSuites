namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with GetVersions operation.
    /// </summary>
    [TestClass]
    public abstract class S09_GetVersions : SharedTestSuiteBase
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
        public void S09_GetVersionsInitialization()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 10004, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetVersions operation.");
            }

            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases for "GetVersions" sub-request.

        /// <summary>
        /// A method used to verify that GetVersions sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S09_TC01_GetVersions_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetVersions"sub-request with correct input parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { getVersionsSubRequest },
                "1", 2, 2, null, null, null, null, null, null, true);
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getVersionsSubResponse, "The object 'getVersionsSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getVersionsSubResponse.ErrorCode, "The object 'getVersionsSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the value of attribute "ErrorCode" in the sub-response equals "Success", then capture MS-FSSHTTP_R2029.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2029,
                         @"[In GetVersions Subrequest][The protocol returns results based on the following conditions:] 
                         Otherwise[the processing of the GetVersions subrequest by the protocol server get the requested versions information successfully], the protocol server sets the error code value to ""Success"" to indicate success in processing the GetVersions subrequest.");

                // Capture the requirement MS-FSSHTTP_R10004
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         10004,
                         @"[In Appendix B: Product Behavior] Implementation does support this operation[GetVersions]. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");

                if (Common.IsRequirementEnabled(11275, this.Site))
                {
                    // Capture the requirement MS-FSSHTTP_R11275
                    Site.CaptureRequirementIfIsNotNull(
                             cellStoreageResponse.ResponseCollection.Response[0].ResourceID,
                             "MS-FSSHTTP",
                             11275,
                             @"[In Appendix B: Product Behavior] The ResourceID attribute is present when the UseResourceID attribute is set to true in the corresponding Request element, [and SHOULD NOT be present otherwise]. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11025
                    Site.CaptureRequirementIfIsNotNull(
                            cellStoreageResponse.ResponseCollection.Response[0].ResourceID,
                            "MS-FSSHTTP",
                            11025,
                            @"[In Response] ResourceID: A string that specifies the invariant ResourceID for a file, which uniquely identifies the file whose response is being generated.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11029
                    Site.CaptureRequirementIfIsTrue(
                            !string.IsNullOrEmpty(cellStoreageResponse.ResponseCollection.Response[0].ResourceID),
                            "MS-FSSHTTP",
                            11029,
                            @"[In Response] [ResourceID] If present, the string value MUST NOT be an empty string.");
                }
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                    @"[In GetVersions Subrequest][The protocol returns results based on the following conditions:] 
                            Otherwise[the processing of the GetVersions subrequest by the protocol server get the requested versions information successfully], the protocol server sets the error code value to ""Success"" to indicate success in processing the GetVersions subrequest.");
            }
        }

        /// <summary>
        /// A method used to verify the value of attribute "List Id" under "Results" element in the sub-response 
        /// when the GetVersion sub-request is executed successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S09_TC02_GetVersions_Success_Results_ListId()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the GUID of expected list using SUT control Adapter method.
            string listName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            string expectedListGuid = SutPowerShellAdapter.GetListGuidByName(listName);

            // Invoke "GetVersions"sub-request with correct input parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { getVersionsSubRequest },
                "1", 2, 2, null, null, null, null, null, null, false);
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getVersionsSubResponse, "The object 'getVersionsSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getVersionsSubResponse.ErrorCode, "The object 'getVersionsSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Make sure the error code value in the response equals "Success"
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                    @"The response of the ""getVersions"" sub-request on the file {0} should succeed.",
                    this.DefaultFileUrl);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2302
                Site.CaptureRequirementIfAreEqual<System.Guid>(
                         new System.Guid(expectedListGuid),
                         new System.Guid(getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.list.id),
                         "MS-FSSHTTP",
                         2302,
                         @"[In GetVersionsSubResponseType][Results complex type] list.id: Specifies the GUID of the document library in which the file resides.");

                if (Common.IsRequirementEnabled(11276, this.Site))
                {
                    // Capture the requirement MS-FSSHTTP_R11276
                    Site.CaptureRequirementIfIsNull(
                             cellStoreageResponse.ResponseCollection.Response[0].ResourceID,
                             "MS-FSSHTTP",
                             11276,
                             @"[In Appendix B: Product Behavior] The ResourceID attribute [MAY be present when the UseResourceID attribute is set to true in the corresponding Request element, and] is not present otherwise[when the UseResourceID attribute is set to false in the corresponding Request element]. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }
            }
            else
            {
                Site.Assert.AreEqual<System.Guid>(
                    new System.Guid(expectedListGuid),
                    new System.Guid(getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.list.id),
                    @"[In GetVersionsSubResponseType][Results complex type] list.id: Specifies the GUID of the document library in which the file resides.");
            }
        }

        /// <summary>
        /// A method used to verify the value of attribute "Versioning Enabled" under "Results" element in the sub-response, 
        /// when the GetVersions sub-request is executed successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S09_TC03_GetVersions_Success_Results_VersioningDiabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, false))
            {
                this.Site.Assert.Fail("Cannot disable the version on the document library {0}", documentLibraryName);
            }

            this.StatusManager.RecordDisableVersioning(documentLibraryName);

            // Invoke GetVersions sub-request with the file URL that under a document list whose versioning is disabled.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getVersionsSubResponse, "The object 'getVersionsSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getVersionsSubResponse.ErrorCode, "The object 'getVersionsSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Make sure the error code value in the response equals "Success".
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                    @"The response of the ""getVersions"" sub-request on the file {0} should succeed.",
                    this.DefaultFileUrl);

                // If the value of "versioning enabled" is 0 under "results" element in the sub-response, then capture MS-FSSHTTP_R2304.
                Site.CaptureRequirementIfAreEqual<byte>(
                         0,
                         getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled,
                         "MS-FSSHTTP",
                         2304,
                         @"[In GetVersionsSubResponseType][Results complex type] versioning.enabled: A value of ""0"" indicates that versioning is disabled.");
            }
            else
            {
                Site.Assert.AreEqual<byte>(
                    0,
                    getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled,
                    @"[In GetVersionsSubResponseType][Results complex type] versioning.enabled: A value of ""0"" indicates that versioning is disabled.");
            }
        }

        /// <summary>
        /// A method used to verify the value of versioning.enabled attribute equals 1 when the versioning of the file is enabled.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S09_TC04_GetVersions_Success_Results_VersioningEnabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot disable the version on the document library {0}", documentLibraryName);
            }

            // Invoke "GetVersions"sub-request with the file URL that under a document list whose versioning is disabled.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getVersionsSubResponse, "The object 'getVersionsSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getVersionsSubResponse.ErrorCode, "The object 'getVersionsSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Make sure the error code value in the response equals "Success".
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                    @"The response of the ""getVersions"" sub-request on the file {0} should succeed.",
                    this.DefaultFileUrl);

                // If the value of "versioning enabled" is 1 under "results" element in the sub-response, then capture MS-FSSHTTP_R2305.
                Site.CaptureRequirementIfAreEqual<byte>(
                         1,
                         getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled,
                         "MS-FSSHTTP",
                         2305,
                         @"[In GetVersionsSubResponseType][Results complex type] versioning.enabled: a value of ""1"" indicates that versioning is enabled. ");
            }
            else
            {
                Site.Assert.AreEqual<byte>(
                    1,
                    getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.versioning.enabled,
                    @"[In GetVersionsSubResponseType][Results complex type] versioning.enabled: a value of ""1"" indicates that versioning is enabled. ");
            }
        }

        /// <summary>
        /// A method used to verify the value of "Version Data" under "Results" element in the sub-response, 
        /// when the GetVersions sub-request is executed successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S09_TC05_GetVersions_Success_Results_VersionData()
        {
            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot disable the version on the document library {0}", documentLibraryName);
            }

            // Prepare a file.
            string fileUrl = this.PrepareFile();

            // Initialize the context
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(fileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", fileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(fileUrl, this.UserName01, this.Password01, this.Domain);

            string checkInComments1 = "New Comment1 for testing purpose on the operation GetVersions.";
            if (!SutPowerShellAdapter.CheckInFile(fileUrl, this.UserName01, this.Password01, this.Domain, checkInComments1))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", fileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.CancelRecordCheckOut(fileUrl);

            // Check out one file by a specified user name again.
            if (!this.SutPowerShellAdapter.CheckOutFile(fileUrl, this.UserName02, this.Password02, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", fileUrl, this.UserName02, this.Password02);
            }

            this.StatusManager.RecordFileCheckOut(fileUrl, this.UserName02, this.Password02, this.Domain);

            string checkInComments2 = "New Comment2 for testing purpose on the operation GetVersions.";
            if (!SutPowerShellAdapter.CheckInFile(fileUrl, this.UserName02, this.Password02, this.Domain, checkInComments2))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", fileUrl, this.UserName02, this.Password02);
            }

            this.StatusManager.CancelRecordCheckOut(fileUrl);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            int contentLength = new IntermediateNodeObject.RootNodeObjectBuilder().Build(
                fsshttpbResponse.DataElementPackage.DataElements,
                fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID).GetContent().Count;

            // Invoke "GetVersions" sub-request with the test file URL that under a document list which is enable versioning.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getVersionsSubResponse, "The object 'getVersionsSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getVersionsSubResponse.ErrorCode, "The object 'getVersionsSubResponse.ErrorCode' should not be null.");

            // Make sure the error code value in the response equals "Success"
            Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                @"The response of the ""getVersions"" sub-request on the file {0} should succeed.",
                this.DefaultFileUrl);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the result contains 3 versions (1.the original, 2. User01 checked in, 3 User02 checked in), then the MS-FSSHTTP requirement: MS-FSSHTTP_R2307 is verified.
                Site.CaptureRequirementIfAreEqual<int>(
                         3,
                         getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.result.Length,
                         "MS-FSSHTTP",
                         2307,
                         @"[In GetVersionsSubResponseType][Results complex type] result: A separate result element MUST exist for each version of the file that the user can access.");

                // If the result contains 3 versions (1.the original, 2. User01 checked in, 3 User02 checked in), then the MS-FSSHTTP requirement: MS-FSSHTTP_R30841 is verified.
                Site.CaptureRequirementIfAreEqual<int>(
                         3,
                         getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.result.Length,
                         "MS-FSSHTTP",
                         30841,
                         @"[In GetVersionsResponse] GetVersionsResult: An XML node that contains the details about all the versions of the specified file that the user can access.");
            }
            else
            {
                Site.Assert.AreEqual<int>(
                    3,
                    getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.result.Length,
                    @"[In GetVersionsSubResponseType][Results complex type] result: A separate result element MUST exist for each version of the file that the user can access.");
            }

            bool isFindVersion = false;
            bool isNotStartWithAt = false;

            foreach (VersionData item in getVersionsSubResponse.GetVersionsResponse.GetVersionsResult.results.result)
            {
                if (string.Compare(checkInComments2, item.comments, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    bool isCreatedBy = item.createdBy != null && item.createdBy.IndexOf(this.UserName02, System.StringComparison.OrdinalIgnoreCase) >= 0;
                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "For requirement MS-FSSHTTP_R2312, expect the CreatedBy contains the user name {0}, the actual value is {1}",
                        this.UserName02,
                        item.createdBy);

                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2312
                        Site.CaptureRequirementIfIsTrue(
                            isCreatedBy,
                            "MS-FSSHTTP",
                            2312,
                            @"[In GetVersionsSubResponseType][VersionData complex type] createdBy: The creator of the version of the file.");

                        bool isCreatedByName = item.createdByName != null && item.createdByName.IndexOf(this.UserName02, System.StringComparison.OrdinalIgnoreCase) >= 0;
                        this.Site.Log.Add(
                            LogEntryKind.Debug,
                            "For requirement MS-FSSHTTP_R2313, expect the createdByName contains the user name {0}, the actual value is {1}",
                            this.UserName02,
                            item.createdByName);

                        // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2313
                        Site.CaptureRequirementIfIsTrue(
                                 isCreatedByName,
                                 "MS-FSSHTTP",
                                 2313,
                                 @"[In GetVersionsSubResponseType][VersionData complex type] createdByName: The display name of the creator of the version of the file.");

                        bool isVersion = item.version != null && item.version.StartsWith("@", StringComparison.OrdinalIgnoreCase);
                        this.Site.Log.Add(
                            LogEntryKind.Debug,
                            "For requirement MS-FSSHTTP_R2309, expect the version start with @, the actual value is {0}",
                            item.version);

                        // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2309
                        Site.CaptureRequirementIfIsTrue(
                                 isVersion,
                                 "MS-FSSHTTP",
                                 2309,
                                 @"[In GetVersionsSubResponseType][VersionData complex type] version: The most recent version of the file MUST be preceded with an at sign (@).");

                        // Using UNICODE, so the actual file size will double the length 5. 
                        Site.CaptureRequirementIfAreEqual<ulong>(
                                 (ulong)contentLength,
                                 item.size,
                                 "MS-FSSHTTP",
                                 2314,
                                 @"[In GetVersionsSubResponseType][VersionData complex type] size: The size, in bytes, of the version of the file.");

                        // If go through here, then the requirement MS-FSSHTTP_R2315 can be directly captured. 
                        Site.CaptureRequirement(
                                 "MS-FSSHTTP",
                                 2315,
                                 @"[In GetVersionsSubResponseType][VersionData complex type] comments: The comment entered when the version of the file was replaced on the protocol server during check in.");
                    }
                    else
                    {
                        Site.Assert.IsTrue(
                            isCreatedBy,
                            @"[In GetVersionsSubResponseType][VersionData complex type] createdBy: The creator of the version of the file.");

                        Site.Assert.IsNotNull(
                            item.createdRaw,
                            @"[In GetVersionsSubResponseType] Implementation does return this attribute[createdRaw]. [In VersionData] createdRaw: The creation date and time for the version of the file in DateTime format, as specified in [ISO-8601]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");

                        bool isCreatedByName = item.createdByName != null && item.createdByName.IndexOf(this.UserName02, System.StringComparison.OrdinalIgnoreCase) >= 0;
                        Site.Assert.IsTrue(
                            isCreatedByName,
                            @"[In GetVersionsSubResponseType][VersionData complex type] createdByName: The display name of the creator of the version of the file.");

                        bool isVersion = item.version != null && item.version.StartsWith("@", StringComparison.OrdinalIgnoreCase);
                        Site.Assert.IsTrue(
                            isVersion,
                            @"[In GetVersionsSubResponseType][VersionData complex type] version: The most recent version of the file MUST be preceded with an at sign (@).");

                        Site.Assert.AreEqual<ulong>(
                            (ulong)contentLength,
                            item.size,
                            @"[In GetVersionsSubResponseType][VersionData complex type] size: The size, in bytes, of the version of the file.");
                    }

                    isFindVersion = true;
                }
                else
                {
                    isNotStartWithAt = !item.version.StartsWith("@", StringComparison.OrdinalIgnoreCase);
                }
            }

            if (!isFindVersion)
            {
                this.Site.Assert.Fail("Cannot find the Version record for the comment {0}", checkInComments2);
            }

            Site.Log.Add(
                LogEntryKind.Debug,
                "All the other versions MUST exist without any prefix, and actually they {0} have prefix.",
                isNotStartWithAt ? "do not" : "do");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2310
                Site.CaptureRequirementIfIsTrue(
                         isNotStartWithAt,
                         "MS-FSSHTTP",
                         2310,
                         @"[In GetVersionsSubResponseType][VersionData complex type] version: All the other versions MUST exist without any prefix. ");
            }
            else
            {
                Site.Assert.IsTrue(
                    isNotStartWithAt,
                    @"[In GetVersionsSubResponseType][VersionData complex type] version: All the other versions MUST exist without any prefix. ");
            }
        }

        #endregion
    }
}