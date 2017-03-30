namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with Versioning operation.
    /// </summary>
    [TestClass]
    public abstract class S16_Versioning : SharedTestSuiteBase
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
        public void S16_VersioningInitialization()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 111341, this.Site), "This test case only runs when Versioning subrequest is supported.");
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "Versioning" sub-request.

        /// <summary>
        /// A method used to verify that Versioning sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC01_Versioning_GetVersionList_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = SharedTestSuiteHelper.CreateGetDocMetaInfoSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getDocMetaInfoSubRequest });
            GetDocMetaInfoSubResponseType getDocMetaInfoSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetDocMetaInfoSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(getDocMetaInfoSubResponse, "The object 'getDocMetaInfoSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(getDocMetaInfoSubResponse.ErrorCode, "The object 'getDocMetaInfoSubResponse.ErrorCode' should not be null.");

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(versioningSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(versioningSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11146
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.UserTable,
                    "MS-FSSHTTP",
                    11146,
                    @"[In VersioningSubResponseDataType] The UserTable element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11053
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.UserTable,
                    "MS-FSSHTTP",
                    11053,
                    @"[In SubResponseDataGenericType] The UserTable element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11148
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.Versions,
                    "MS-FSSHTTP",
                    11148,
                    @"[In VersioningSubResponseDataType] The Versions element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11055
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.Versions,
                    "MS-FSSHTTP",
                    11055,
                    @"[In SubResponseDataGenericType] The Versions element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                bool isR1117501Verified = bool.Parse(versioningSubResponse.SubResponseData.Versions.Version[0].IsCurrent);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1117501
                Site.CaptureRequirementIfIsTrue(
                    isR1117501Verified,
                    "MS-FSSHTTP",
                    1117501,
                    @"[In FileVersionDataType] IsCurrent: True specifies this version is the most recent version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11150
                // This requirement can be captured directly after capturing MS-FSSHTTP_R11146 and MS-FSSHTTP_R11148
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11150,
                    @"[In VersioningSubResponseType] In the case of success, it contains information requested as part of a versioning subrequest.");

                GetDocMetaInfoPropertyType lastModifiedProperty = null;
                foreach (GetDocMetaInfoPropertyType property in getDocMetaInfoSubResponse.SubResponseData.DocProps.Property)
                {
                    if (property.Key.ToLower().Contains("timelastmodified"))
                    {
                        lastModifiedProperty = property;
                    }
                }

                Site.Assert.IsNotNull(lastModifiedProperty, "Property for last modified time should be found.");

                System.DateTime time = Convert.ToDateTime(lastModifiedProperty.Value);

                long lastModifiedTime = long.Parse(versioningSubResponse.SubResponseData.Versions.Version[0].LastModifiedTime);
                bool isR11179Verified = ((lastModifiedTime / 10000000) == (time - new System.DateTime(1601, 1, 1, 0, 0, 0)).TotalSeconds);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11179
                Site.CaptureRequirementIfIsTrue(
                    isR11179Verified,
                    "MS-FSSHTTP",
                    11179,
                    @"[In FileVersionDataType] LastModifiedTime specifies the number of 100-nanosecond intervals that have elapsed since 00:00:00 on January 1, 1601, which MUST be Coordinated Universal Time (UTC).");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11178
                // Calculate expected last modified time by multiple the time span with ten million, this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isR11179Verified,
                    "MS-FSSHTTP",
                    11178,
                    @"[In FileVersionDataType] A single tick represents 100 nanoseconds, or one ten-millionth of a second.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11180
                Site.CaptureRequirementIfAreEqual<string>(
                    versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                    versioningSubResponse.SubResponseData.Versions.Version[0].UserId,
                    "MS-FSSHTTP",
                    11180,
                    @"[In FileVersionDataType] UserId: An integer that specifies the user that last modified the version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11181
                Site.CaptureRequirementIfAreEqual<string>(
                    versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                    versioningSubResponse.SubResponseData.Versions.Version[0].UserId,
                    "MS-FSSHTTP",
                    11181,
                    @"[In FileVersionDataType] The number MUST match the UserId attribute of a UserDataType (section 2.3.1.42) described in the VersioningUserTableType in the current VersioningSubResponseDataType.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11249
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11249,
                    @"[In Get Version List] If the VersioningRequestType attribute is set to ""GetVersionList"", the protocol server considers the versioning subrequest to be of type ""Get Version List"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11250
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11250,
                    @"[In Get Version List] The protocol server processes this request to return a list of the most recent versions of the file.");
            }
            else
            {
                Site.Assert.IsNotNull(
                    versioningSubResponse.SubResponseData.UserTable,
                    @"The UserTable element must be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                Site.Assert.IsNotNull(
                    versioningSubResponse.SubResponseData.Versions,
                    @"The UserTable element must be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");
            }
        }

        /// <summary>
        /// A method used to verify that restore version can be executed successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC02_Versioning_RestoreVersion_Success()
        {
            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot enable the version on the document library {0}", documentLibraryName);
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string checkInComments1 = "New Comment1 for testing purpose on the operation Versioning.";
            if (!SutPowerShellAdapter.CheckInFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain, checkInComments1))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.CancelRecordCheckOut(this.DefaultFileUrl);

            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.RestoreVersion, "1.0", this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(versioningSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(versioningSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11248
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    11248,
                    @"[In Versioning Subrequest] [The protocol returns results based on the following conditions:]An ErrorCode value of ""Success"" indicates success in processing the versioning request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11254
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11254,
                    @"[In Restore Version] If the VersioningRequestType attribute is set to ""RestoreVersion"", the protocol server considers the versioning subrequest to be of type ""Restore Version"". ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11255
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11255,
                    @"[In Restore Version] The protocol server processes this request by restoring the file to its state in the version specified by the VersionNumber attribute.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    @"Restore version should succeed.");
            }
        }

        /// <summary>
        /// A method used to verify that error VersionNotFound should be returned if the version is not found.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC03_Versioning_RestoreVersion_VersionNotFound()
        {
            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot enable the version on the document library {0}", documentLibraryName);
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string checkInComments1 = "New Comment1 for testing purpose on the operation Versioning.";
            if (!SutPowerShellAdapter.CheckInFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain, checkInComments1))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.CancelRecordCheckOut(this.DefaultFileUrl);

            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.RestoreVersion, "3.0", this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(versioningSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(versioningSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11247
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.VersionNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    11247,
                    @"[In Versioning Subrequest] [The protocol returns results based on the following conditions:]If the protocol server gets a versioning subrequest of type ""Restore version"" and the restore fails because the version number specifies a non-existent version, the protocol server returns an error code value set to ""VersionNotFound"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11256
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.VersionNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    11256,
                    @"[In Restore Version] If the VersionNumber attribute specifies a version that doesn't exist, the protocol server returns an error status set to ""VersionNotFound"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11083
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.VersionNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    11083,
                    @"[In VersioningRelatedErrorCodeTypes] The value ""VersionNotFound"" indicates that the version number specified by the protocol client doesn’t match a version of the file on the protocol server.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.VersionNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    @"Error VersionNotFound should be returned if restore version with not found version.");
            }
        }

        /// <summary>
        /// A method used to verify that FileVersionEventDataType with type 3 indicates a user restored the file content to its state at a previous version.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC04_FileVersionEventDataType_Restore()
        {
            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot enable the version on the document library {0}", documentLibraryName);
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string checkInComments1 = "New Comment1 for testing purpose on the operation Versioning.";
            if (!SutPowerShellAdapter.CheckInFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain, checkInComments1))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.CancelRecordCheckOut(this.DefaultFileUrl);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.RestoreVersion, "1.0", this.Site);
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                "Restore file version should succeed.");

            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = SharedTestSuiteHelper.CreateGetDocMetaInfoSubRequest(SequenceNumberGenerator.GetCurrentToken());
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getDocMetaInfoSubRequest });
            GetDocMetaInfoSubResponseType getDocMetaInfoSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetDocMetaInfoSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(getDocMetaInfoSubResponse.ErrorCode, this.Site),
                "Get doc meta info should succeed.");

            versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { versioningSubRequest });
            versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                "Get version list should succeed.");
            this.Site.Assert.AreEqual<int>(
                3,
                versioningSubResponse.SubResponseData.Versions.Version.Length,
                "There should be 3 version numbers.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11190
            Site.CaptureRequirementIfAreEqual<string>(
                "3",
                versioningSubResponse.SubResponseData.Versions.Version[0].Events.Event[0].Type,
                "MS-FSSHTTP",
                11190,
                @"[In FileVersionEventDataType] 3 means A user restored the file content to its state at a previous version.");

            GetDocMetaInfoPropertyType lastModifiedProperty = null;
            foreach (GetDocMetaInfoPropertyType property in getDocMetaInfoSubResponse.SubResponseData.DocProps.Property)
            {
                if (property.Key.ToLower().Contains("timelastmodified"))
                {
                    lastModifiedProperty = property;
                }
            }

            Site.Assert.IsNotNull(lastModifiedProperty, "Property for last modified time should be found.");

            System.DateTime time = Convert.ToDateTime(lastModifiedProperty.Value);

            long createTime = long.Parse(versioningSubResponse.SubResponseData.Versions.Version[0].Events.Event[0].CreateTime);
            bool isR11193Verified = ((System.Math.Round((double)createTime / 10000000)) == (time - new System.DateTime(1601, 1, 1, 0, 0, 0)).TotalSeconds);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11193
            Site.CaptureRequirementIfIsTrue(
                isR11193Verified,
                "MS-FSSHTTP",
                11193,
                @"[In FileVersionEventDataType] CreateTime specifies the number of 100-nanosecond intervals that have elapsed since 00:00:00 on January 1, 1601, which MUST be Coordinated Universal Time (UTC).");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11192
            Site.CaptureRequirementIfIsTrue(
                isR11193Verified,
                "MS-FSSHTTP",
                11192,
                @"[In FileVersionEventDataType] A single tick represents 100 nanoseconds, or one ten-millionth of a second.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11194
            Site.CaptureRequirementIfAreEqual<string>(
                versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                versioningSubResponse.SubResponseData.Versions.Version[0].Events.Event[0].UserId,
                "MS-FSSHTTP",
                11194,
                @"[In FileVersionEventDataType] UserId: An integer which specifies which user performed this event.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11195
            Site.CaptureRequirementIfAreEqual<string>(
                versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                versioningSubResponse.SubResponseData.Versions.Version[0].Events.Event[0].UserId,
                "MS-FSSHTTP",
                11195,
                @"[In FileVersionEventDataType] The UserId MUST match the UserId attribute of a UserDataType (section 2.3.1.42) described in the VersioningUserTableType in the current VersioningSubResponseDataType.");
        }

        /// <summary>
        /// A method used to verify that FileOperation sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC05_FileVersionEventDataType_Rename()
        {
            string documentLibraryName = Common.GetConfigurationPropertyValue("MSFSSHTTPFSSHTTPBLibraryName", this.Site);
            if (!SutPowerShellAdapter.SwitchMajorVersioning(documentLibraryName, true))
            {
                this.Site.Assert.Fail("Cannot enable the version on the document library {0}", documentLibraryName);
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string checkInComments1 = "New Comment1 for testing purpose on the operation Versioning.";
            if (!SutPowerShellAdapter.CheckInFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain, checkInComments1))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check in status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.CancelRecordCheckOut(this.DefaultFileUrl);

            string fileName = this.DefaultFileUrl.Substring(this.DefaultFileUrl.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);
            string newName = Common.GenerateResourceName(this.Site, "fileName") + ".txt";

            FileOperationSubRequestType fileOperationSubRequest = SharedTestSuiteHelper.CreateFileOperationSubRequest(FileOperationRequestTypes.Rename, newName, null, this.Site);

            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { fileOperationSubRequest });

            FileOperationSubResponseType fileOperationSubResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(fileOperationSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(fileOperationSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

            string fileUrl = this.DefaultFileUrl.Substring(0, this.DefaultFileUrl.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1) + newName;

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                "Get version list should succeed.");
            this.Site.Assert.AreEqual<int>(
                2,
                versioningSubResponse.SubResponseData.Versions.Version.Length,
                "There should be 2 version numbers.");

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11189
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                versioningSubResponse.SubResponseData.Versions.Version[0].Events.Event[0].Type,
                "MS-FSSHTTP",
                11189,
                @"[In FileVersionEventDataType] 2 means A user renamed the file.");
        }
        #endregion
    }
}