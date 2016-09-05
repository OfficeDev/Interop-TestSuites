namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Linq;

    /// <summary>
    /// This test class is used to test GetRoutingDestinationUrl and SubmitFile operation.
    /// </summary>
    [TestClass]
    public class S01_GetRoutingDestinationUrlAndSubmitFile : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is not configured for routing content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC01_GetFinalRoutingDestinationFolderUrl_NotConfiguredForRouting()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use the repository that is not configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.DisableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is not configured for routing content, expect the server responses DocumentRoutingDisabled.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.DefaultLibraryUrl);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R64 and MS-OFFICIALFILE_R124
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.DocumentRoutingDisabled,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     64,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 1. If the repository is not configured for routing content then the protocol server MUST set the ResultType element to ""DocumentRoutingDisabled"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.DocumentRoutingDisabled,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     124,
                     @"[In DocumentRoutingResultType] [If] Routing is not enabled on the repository, [DocumentRoutingResultType] value is DocumentRoutingDisabled.");
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is not configured for routing content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC02_SubmitFile_NotConfiguredForRouting()
        {
            if (!Common.Common.IsRequirementEnabled(201, this.Site))
            {
                this.Site.Assume.Inconclusive("Implementation cannot be configured to disable the record routing.");
            }

            // Initial parameters to use the repository that is not configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.DisableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file Property.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // Convert the binary input into Base64 UUEncoded output.
            byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

            // Submit a file and its associated properties to the repository which is not configured for content routing, expect the server responses InvalidRouterConfiguration.
            SubmitFileResult submitFile = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, this.DefaultLibraryUrl, this.SubmitUserName);

            // We have deactivated routing of the repository to which files are submitted,
            // so the ResultCode should be InvalidRouterConfiguration as specified of TD. 
            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R201
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.InvalidRouterConfiguration,
                     submitFile.ResultCode,
                     "MS-OFFICIALFILE",
                     201,
                     @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] If the protocol server determines that the repository is not configured for routing, the protocol server MUST set the ResultCode element to ""InvalidRouterConfiguration"" and return.  (Microsoft SharePoint Server 2010 and above follow this behavior.)");

            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.InvalidRouterConfiguration,
                     submitFile.ResultCode,
                     "MS-OFFICIALFILE",
                     250,
                     @"[In SubmitFileResultCode] [If] The operation [SubmitFile] failed because the repository was not configured for routing. [SubmitFileResultCode] Value is InvalidRouterConfiguration.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is a storage location where the user has not permissions to store content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC03_GetFinalRoutingDestinationFolderUrl_UserNoPermissions()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use the repository that is a storage location where the user has not permissions to store content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.LimitedUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.LimitedUserPassword;
            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is a storage location where the user has not permissions to store content, expect the server responses PermissionDeniedAtDestination.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // This requirement will be partially captured, the Url to the temporary storage location will not be tested due to the temporary location is server internal behavior. 
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.PermissionDeniedAtDestination,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     71,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 8. If the user does not have permissions to store content in the storage location determined by the rules, then the protocol server MUST set ResultType to ""PermissionDeniedAtDestination"" and Url to the temporary storage location and return.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R125
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.PermissionDeniedAtDestination,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     125,
                     @"[In DocumentRoutingResultType] [If] Implementation-specific error validating permissions for the user at the storage location, [DocumentRoutingResultType] value is PermissionDeniedAtDestination.");
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is a storage location where the user has not permissions to store content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC04_SubmitFile_UserNoPermissions()
        {
            // Initial parameters to use the repository that is a storage location where the user has not permissions to store content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.LimitedUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.LimitedUserPassword;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // Convert the binary input into Base64 UUEncoded output.
            byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

            // Submit a file and its associated properties to the repository that is a storage location where the user has not permissions to store content, expect the server responses NotFound error code.
            SubmitFileResult submitFileResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, this.DocumentLibraryUrlOfAppendUniqueSuffix, this.LimitedUserName);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R199
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.NotFound,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     199,
                     @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] If the user is not in the list of users authorized to submit files to the repository, the protocol server MUST set the ResultCode element to ""NotFound"" and return.");
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.NotFound,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     253,
                     @"[In SubmitFileResultCode] [If] The operation [SubmitFile] failed because the user was not authorized to submit files. [SubmitFileResultCode] Value is NotFound.");
        }

        /// <summary>
        /// This test case is used to test the GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content when the required properties are missed.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC05_GetFinalRoutingDestinationFolderUrl_MissingRequiredProperties()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which will not contain all required properties.
            RecordsRepositoryProperty[] partialProperties = this.ConstructPartialRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content when the required properties are missed, expect the server responses SuccessToDropOffLibrary.
            DocumentRoutingResult docRoutingResult = this.Adapter.GetFinalRoutingDestinationFolderUrl(partialProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R70
            // This requirement will be partially captured, the Url to the temporary storage location will not be tested due to the temporary location is server internal behavior.
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.SuccessToDropOffLibrary,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     70,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 7. If the protocol server determines that the storage location determined by the rules has required properties that are not present in the properties element, the protocol server MUST set the ResultType element to ""SuccessToDropOffLibrary"", set the Url to the temporary storage location and return.");
        }

        /// <summary>
        /// This test case is used to test the GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content when the OriginalSaveLocation parameter is empty string.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC06_GetFinalRoutingDestinationFolderUrl_EmptyOriginalSaveLocation()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content when the OriginalSaveLocation parameter is empty string, expect the server responses UnknownFailure.
            DocumentRoutingResult docRoutingResult = this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, string.Empty);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R7201 and MS-OFFICIALFILE_R129
            // We verify this requirement partially.
            // We can get UnknownFailure when use a empty originalSaveLocation. 
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.UnknownFailure,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     7201,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 9. If we use a empty originalSaveLocation, the MS product set the ResultType element to ""UnknownFailure"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.UnknownFailure,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     129,
                     @"[In DocumentRoutingResultType] [If] Generic error, [DocumentRoutingResultType] value is UnknownFailure.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and parsing file metadata is disabled.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC07_GetFinalRoutingDestinationFolderUrl_ParsingDisabledAtDestination()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial server parameters to use a repository that is configured for routing content and parsing file metadata is disabled.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            bool isSucceed = false;
            string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);

            // Turn off the file metadata parsing on the record center site.
            isSucceed = this.SutControlAdapter.SwitchFileMetaDataParsingFeature(siteUrl, false);
            this.Site.Assert.IsTrue(
                isSucceed,
                string.Format("Turning off file metadata parsing should succeed on the site {0}, the actual result is {1}", siteUrl, isSucceed ? "success" : "failure"));

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and parsing file metadata is disabled, expect the server responses ParsingDisabledAtDestination.
            DocumentRoutingResult docRoutingResult = this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            bool isR73Verified = !string.IsNullOrEmpty(docRoutingResult.Url);

            // Add the log information.
            Site.Log.Add(
                    Microsoft.Protocols.TestTools.LogEntryKind.Comment,
                    "Url = {0}, when the GetFinalRoutingDestinationFolderUrl succeeds.", 
                    docRoutingResult.Url);

            Site.Log.Add(
                  Microsoft.Protocols.TestTools.LogEntryKind.Debug,
                  "Expect the URL is not empty when the GetFinalRoutingDestinationFolderUrl succeeds, actually it is {0}",
                  isR73Verified ? "not empty" : "empty");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R73
            Site.CaptureRequirementIfIsTrue(
                     isR73Verified,
                     "MS-OFFICIALFILE",
                     73,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 10. [If other implementation-specific errors occur while determining the storage location for the submission, the protocol server MUST set the ResultType element to ""UnknownFailure"" and return.] Otherwise, the protocol server sets Url to the storage location determined by the rules, which MUST be non-empty.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R92 and MS-OFFICIALFILE_R126
            // This repository is a record center and its default setting is parsing disabled. 
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.ParsingDisabledAtDestination,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     92,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 11. If parsing of file metadata is disabled on the repository, the protocol server MUST set the ResultType to ""ParsingDisabledAtDestination"".");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.ParsingDisabledAtDestination,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     126,
                     @"[In DocumentRoutingResultType] [If] Parsing is not enabled on the repository, [DocumentRoutingResultType] value is ParsingDisabledAtDestination.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and a file with the same name does not already exist at the location determined by the rule.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC08_GetFinalRoutingDestinationFolderUrl_NoCollision()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and a file with the same name does not already exist at the location determined by the rule, expect the server responses DocumentRoutingCollisionSetting as NoCollision.
            DocumentRoutingResult docRoutingResult = this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R75 and MS-OFFICIALFILE_R132
            // In this situation, we did not submit any same file to this location and we 
            // did not configure any file in environment setup, so we know this.
            Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                     DocumentRoutingCollisionSetting.NoCollision,
                     docRoutingResult.CollisionSetting,
                     "MS-OFFICIALFILE",
                     75,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 12. If the ResultType is set to ""Success"", ""SuccessToDropOffLibrary"" or ""PermissionDeniedAtDestination"": When a file with the same name does not already exist at the location determined by the rule then the protocol server MUST set CollisionSetting to ""NoCollision"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                     DocumentRoutingCollisionSetting.NoCollision,
                     docRoutingResult.CollisionSetting,
                     "MS-OFFICIALFILE",
                     132,
                     @"[In DocumentRoutingCollisionSetting] [If] There was no file of the same name already stored at the storage location, [DocumentRoutingCollisionSetting] value is NoCollision.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and using "SharePointVersioning" when there is a file with the same name at the location determined by the rule.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC09_GetFinalRoutingDestinationFolderUrl_UseSharePointVersioningService()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureDocumentsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            try
            {
                // Properties of the file to submit which is configured in the file PropertyConfig.xml.
                RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

                // Convert the binary input into Base64 UUEncoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));
                string fileName = this.GenerateRandomTextFileName();

                // Submit a file and its associated properties to the repository on a repository that is configured for routing content and using "SharePointVersioning" when there is a file with the same name at the location determined by the rule.
                SubmitFileResult submitResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, fileName, this.SubmitUserName);

                this.Site.Assert.AreEqual<SubmitFileResultCode?>(
                    SubmitFileResultCode.Success,
                    submitResult.ResultCode,
                    "Submit the file with the name {0} should succeed.",
                    fileName);

                // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and using "SharePointVersioning" with the same file name as the previous step submitted, expect the server responses DocumentRoutingCollisionSetting as UseSharePointVersioning.
                DocumentRoutingResult docRoutingResult =
                    this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, submitResult.ResultUrl);

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R366 and MS-OFFICIALFILE_R133
                Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                         DocumentRoutingCollisionSetting.UseSharePointVersioning,
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         366,
                         @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:][12. If the ResultType is set to ""Success"", ""SuccessToDropOffLibrary"" or ""PermissionDeniedAtDestination"": When  a file with the same name does not already exist at the location determined by the rule then the protocol server MUST set CollisionSetting to ""NoCollision"" and return;]  when a file with the same name does exist, if the repository is configured to overwrite existing files and versioning is enabled at the location determined by the rule, then the protocol server MUST set CollisionSetting to ""UseSharePointVersioning"" and return.");

                Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                         DocumentRoutingCollisionSetting.UseSharePointVersioning,
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         133,
                         @"[In DocumentRoutingCollisionSetting] [If] Overwrite the existing file, [DocumentRoutingCollisionSetting] value is UseSharePointVersioning.");
            }
            finally
            {
                string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerDocumentsCenterSite", this.Site);
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DocumentLibraryName", this.Site);
                bool isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);

                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actually result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));
            }
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and using "AppendUniqueSuffixes" when there is a file with the same name at the location determined by the rule.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC10_GetFinalRoutingDestinationFolderUrl_AppendUniqueSuffixes()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = Common.Common.GetConfigurationPropertyValue("EnableRoutingFeatureRecordsCenterServiceUrl", this.Site);
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            bool isEnableFileParsingSucceed = false;
            string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);
            try
            {
                // Turn on the file metadata parsing on the record center site.
                isEnableFileParsingSucceed = this.SutControlAdapter.SwitchFileMetaDataParsingFeature(siteUrl, true);
                this.Site.Assert.IsTrue(
                    isEnableFileParsingSucceed,
                    string.Format("Turning on file metadata parsing should succeed on the site {0}, the actual result is {1}", siteUrl, isEnableFileParsingSucceed ? "success" : "failure"));

                // Properties of the file to submit which is configured in the file PropertyConfig.xml.
                RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

                // Convert the binary input into Base64 UUEncoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));
                string fileName = this.GenerateRandomTextFileName();

                // Submit a file and its associated properties to the repository.
                SubmitFileResult submitFileResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, fileName, this.SubmitUserName);
                this.Site.Assert.AreEqual(
                    SubmitFileResultCode.Success,
                    submitFileResult.ResultCode,
                    "Test case cannot continue if the submit the file {0} to the library fails.",
                    fileName,
                    this.DocumentLibraryUrlOfAppendUniqueSuffix);

                // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and using "AppendUniqueSuffixes" with the same file name as the previous step submitted, expect the server responds DocumentRoutingCollisionSetting as AppendUniqueSuffixes.
                DocumentRoutingResult docRoutingResult = this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix, fileName));

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R77
                Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                         DocumentRoutingCollisionSetting.AppendUniqueSuffixes,
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         77,
                         @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:][12. If the ResultType is set to ""Success"", ""SuccessToDropOffLibrary"" or ""PermissionDeniedAtDestination"": When a file with the same name does not already exist at the location determined by the rule, then the protocol server MUST set CollisionSetting to ""NoCollision"" and return; when a file with the same name does exist, if the repository is configured to overwrite existing files and versioning is enabled at the location determined by the rule, then the protocol server MUST set CollisionSetting to ""UseSharePointVersioning"" and return, ] otherwise the protocol server MUST set CollisionSetting to ""AppendUniqueSuffixes"".");

                Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                         DocumentRoutingCollisionSetting.AppendUniqueSuffixes,
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         117,
                         @"[DocumentRoutingResult] CollisionSetting: Collision setting, that is, whether there was a file of the same name already stored at the storage location and, if so, the server configured behavior.");

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R134
                Site.CaptureRequirementIfAreEqual<DocumentRoutingCollisionSetting>(
                         DocumentRoutingCollisionSetting.AppendUniqueSuffixes,
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         134,
                         @"[In DocumentRoutingCollisionSetting] [If] Generate a unique file name, [DocumentRoutingCollisionSetting] value is AppendUniqueSuffixes.");

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R130002
                Site.CaptureRequirementIfIsNotNull(
                         docRoutingResult.CollisionSetting,
                         "MS-OFFICIALFILE",
                         130002,
                         @"[In DocumentRoutingCollisionSetting] Collision setting for the repository, that is, whether there was a file of the same name already stored at the storage location and, if so, the server configured behavior.");
            }
            finally
            {
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DocumentLibraryName", this.Site);
                bool isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);

                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actual result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));

                if (isEnableFileParsingSucceed)
                {
                    // Turn off the file metadata parsing on the record center site.
                    bool isDisableFileParsingSucceed = this.SutControlAdapter.SwitchFileMetaDataParsingFeature(siteUrl, false);
                    this.Site.Assume.IsTrue(
                        isDisableFileParsingSucceed,
                        string.Format("Turning off file metadata parsing should succeed on the site {0}, the actual result is {1}", siteUrl, isEnableFileParsingSucceed ? "success" : "failure"));
                }
            }
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and parsing file metadata is enabled.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC11_GetFinalRoutingDestinationFolderUrl_Success()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureDocumentsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and parsing file metadata is enabled.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfSharePointVersion));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R93
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.Success,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     93,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] [11. If parsing of file metadata is disabled on the repository, the protocol server MUST set the ResultType to ""ParsingDisabledAtDestination"".] Otherwise the protocol server MUST set the ResultType to ""Success"".");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R120
            Site.CaptureRequirementIfAreEqual(
                     DocumentRoutingResultType.Success,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     120,
                     @"[In DocumentRoutingResultType] [If] The operation [GetFinalRoutingDestinationFolderUrl] was successful, [DocumentRoutingResultType] value is Success.");

            bool isR73Verified = !string.IsNullOrEmpty(docRoutingResult.Url);

            // Add the log information.
            Site.Log.Add(
                    Microsoft.Protocols.TestTools.LogEntryKind.Comment,
                    "Url = {0}, when the GetFinalRoutingDestinationFolderUrl succeeds.", 
                    docRoutingResult.Url);

            Site.Log.Add(
                  Microsoft.Protocols.TestTools.LogEntryKind.Debug,
                  "Expect the URL is not empty when the GetFinalRoutingDestinationFolderUrl succeeds, actually it is {0}",
                  isR73Verified ? "not empty" : "empty");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R73
            Site.CaptureRequirementIfIsTrue(
                     isR73Verified,
                     "MS-OFFICIALFILE",
                     73,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 10. [If other implementation-specific errors occur while determining the storage location for the submission, the protocol server MUST set the ResultType element to ""UnknownFailure"" and return.] Otherwise, the protocol server sets Url to the storage location determined by the rules, which MUST be non-empty.");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R353                    
            Site.CaptureRequirementIfIsNotNull(
                     docRoutingResult,
                     "MS-OFFICIALFILE",
                     353,
                     @"[In Appendix B: Product Behavior] Implementation does provide this method [GetFinalRoutingDestinationFolderUrl]. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the URL of OriginalSaveLocation is a document set.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC12_GetFinalRoutingDestinationFolderUrl_OriginalSaveLocationIsDocumentSet()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial server parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = Common.Common.GetConfigurationPropertyValue("EnableRoutingFeatureRecordsCenterServiceUrl", this.Site);
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the URL of OriginalSaveLocation is a document set, expect the server responses OriginalSaveLocationIsDocumentSet.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentSetUrl));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R65 and MS-OFFICIALFILE_R127
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.OriginalSaveLocationIsDocumentSet,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     65,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 2. If the originalSaveLocation is a storage location for a document set, then the protocol server MUST set the ResultType element to ""OriginalSaveLocationIsDocumentSet"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.OriginalSaveLocationIsDocumentSet,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     127,
                     @"[In DocumentRoutingResultType] [If] The originalSaveLocation is a storage location for storing document sets, [DocumentRoutingResultType] value is OriginalSaveLocationIsDocumentSet.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the OriginalSaveLocation is the location which no routing rule is mapped. 
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC13_GetFinalRoutingDestinationFolderUrl_NoEnforcementAtOriginalSaveLocation()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = Common.Common.GetConfigurationPropertyValue("EnableRoutingFeatureRecordsCenterServiceUrl", this.Site);
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the OriginalSaveLocation is the location which no routing rule is mapped, expect the server responses NoEnforcementAtOriginalSaveLocation.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.NoEnforceLibraryUrl));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R66 and MS-OFFICIALFILE_R128
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.NoEnforcementAtOriginalSaveLocation,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     66,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 3. If the originalSaveLocation is a storage location for which routing is not enforced, then the protocol server MUST set the ResultType element to ""NoEnforcementAtOriginalSaveLocation"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.NoEnforcementAtOriginalSaveLocation,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     128,
                     @"[In DocumentRoutingResultType] [If] The originalSaveLocation is a storage location for which routing is not enforced, [DocumentRoutingResultType] value is NoEnforcementAtOriginalSaveLocation.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the content type name parameter is empty string.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC14_GetFinalRoutingDestinationFolderUrl_EmptyContentTypeName()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the content type name parameter is empty string, expect the server responses MissingRequiredProperties.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, string.Empty, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R67 and MS-OFFICIALFILE_R122
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.MissingRequiredProperties,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     67,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 4. If the contentTypeName is [NULL or] empty then the protocol server MUST set the ResultType element to ""MissingRequiredProperties"" and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.MissingRequiredProperties,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     122,
                     @"[In DocumentRoutingResultType] [If] Additional data required from the protocol client, [DocumentRoutingResultType] value is MissingRequiredProperties.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the content type name parameter is NULL string.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC15_GetFinalRoutingDestinationFolderUrl_NullContentTypeName()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and the content type name parameter is NULL string, expect the server responses MissingRequiredProperties.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, null, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R67001
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.MissingRequiredProperties,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     67001,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 4. If the contentTypeName is NULL [or empty] then the protocol server MUST set the ResultType element to ""MissingRequiredProperties"" and return.");
        }

        /// <summary>
        /// This test case is used to test GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and no configured routing rule for the specified content type name parameter.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC16_GetFinalRoutingDestinationFolderUrl_NoApplicableRule()
        {
            if (!Common.Common.IsRequirementEnabled(353, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetFinalRoutingDestinationFolderUrl operations.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Properties of the file to submit which is configured in the file PropertyConfig.xml.
            RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

            // call GetFinalRoutingDestinationFolderUrl on a repository that is configured for routing content and no configured routing rule for the specified content type name parameter.
            DocumentRoutingResult docRoutingResult =
                this.Adapter.GetFinalRoutingDestinationFolderUrl(fileProperties, this.NotSupportedContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix));

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R69 and MS-OFFICIALFILE_R121
            // This requirement will be partially captured, the Url to the temporary storage location will not be tested due to the temporary location is server internal behavior.
            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.SuccessToDropOffLibrary,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     69,
                     @"[In GetFinalRoutingDestinationFolderUrl] [The protocol client sends a GetFinalRoutingDestinationFolderUrlSoapIn request WSDL message, and the protocol server MUST respond with a GetFinalRoutingDestinationFolderUrlSoapOut response WSDL message, as follows:] 6. If no applicable rule is found then the protocol server MUST set the ResultType element to ""SuccessToDropOffLibrary"" and set the Url to the temporary storage location and return.");

            Site.CaptureRequirementIfAreEqual<DocumentRoutingResultType>(
                     DocumentRoutingResultType.SuccessToDropOffLibrary,
                     docRoutingResult.ResultType,
                     "MS-OFFICIALFILE",
                     121,
                     @"[In DocumentRoutingResultType] [If] The operation [GetFinalRoutingDestinationFolderUrl] was successful and the location determined by the protocol server is the temporary storage location, [DocumentRoutingResultType] value is SuccessToDropOffLibrary.");
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is a storage location configured for routing content and none of the common properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery and _dlc_hold_searchcontexturl are present.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC17_SubmitFile_Success_NoCommonProperties()
        {
            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            try
            {
                // Convert the binary input into Base64 UnEncoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

                // Properties of the file to submit which is configured in the file PropertyConfig.xml.
                RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

                string randomFileName = this.GenerateRandomTextFileName();

                // Call SubmitFile on a repository that is configured for routing content and none of the common properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery and _dlc_hold_searchcontexturl are present and valid.
                SubmitFileResult submitFileResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, randomFileName, this.SubmitUserName);

                this.VerifySubmitFileSucceed(submitFileResult);

                // If the parent node CustomProcessingResult is null, it can prove that the HoldProcessingResult is also not returned.
                bool isR217Verified = submitFileResult.CustomProcessingResult == null || submitFileResult.CustomProcessingResult.HoldProcessingResult == null;
                Site.Assert.IsTrue(
                        isR217Verified,
                        "For the requirement MS-OFFICIALFILE_R217, expect the HoldProcessingResult is not returned, actual it {0}",
                        isR217Verified ? "does not return" : "returns");

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R217
                Site.CaptureRequirementIfIsTrue(
                         isR217Verified,
                         "MS-OFFICIALFILE",
                         217,
                         @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] If none of the properties [_dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery, _dlc_hold_searchcontexturl] mentioned previously are contained in the properties element then the protocol server MUST NOT set CustomProcessingResult.HoldsProcessingResult.");
            }
            finally
            {
                string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DocumentLibraryName", this.Site);
                bool isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);

                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actually result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));
            }
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is a storage location configured for routing content when the required properties are missed.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC18_SubmitFile_MissingProperties()
        {
            if (!Common.Common.IsRequirementEnabled(2081, this.Site) && !Common.Common.IsRequirementEnabled(2082, this.Site))
            {
                this.Site.Assume.Inconclusive("In the test case MSOFFICIALFILE_S01_TC18_SubmitFile_MissingProperties, at least one of the requirements MS-OFFCIIALFILE_R2081 and MS-OFFICIALFILE_R2082 needs to be enable.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            try
            {
                // Convert the random file content into Base64Encoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

                // Properties of the file to submit which will not contain all required properties.
                RecordsRepositoryProperty[] fileProperties = this.ConstructPartialRequiredProperties();

                string randomFileName = this.GenerateRandomTextFileName();

                // call SubmitFile on a repository that is a storage location configured for routing content when the required properties are missed, expect the server responses MoreInformation.
                SubmitFileResult submitFileResult1 = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, randomFileName, this.SubmitUserName);

                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R208
                if (Common.Common.IsRequirementEnabled(2081, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                             SubmitFileResultCode.MoreInformation,
                             submitFileResult1.ResultCode,
                             "MS-OFFICIALFILE",
                             2081,
                             @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] [If the protocol server determines that the storage location determined by the rules has required properties that are not present in the properties element: 1. If the protocol server determines that the name of the user specified in the userName element is invalid using an implementation-specific validation algorithm, then the protocol server MUST set the ResultCode element to InvalidUser and return.] 2. Otherwise, the protocol server MUST set the ResultCode element to MoreInformation [and the ResultUrl element to an implementation-specific URL to enter more information about the submission.](Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                if (Common.Common.IsRequirementEnabled(2082, this.Site))
                {
                    Site.CaptureRequirementIfIsNotNull(
                             submitFileResult1.ResultUrl,
                             "MS-OFFICIALFILE",
                             2082,
                             @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] [If the protocol server determines that the storage location determined by the rules has required properties that are not present in the properties element: 1. If the protocol server determines that the name of the user specified in the userName element is invalid using an implementation-specific validation algorithm, then the protocol server MUST set the ResultCode element to InvalidUser and return.] 2. Otherwise, the protocol server MUST set[the ResultCode element to MoreInformation and] the ResultUrl element to an implementation-specific URL to enter more information about the submission. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }
            }
            finally
            {
                string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DefaultLibraryName", this.Site);
                string holdName = Common.Common.GetConfigurationPropertyValue("HoldName", this.Site);

                bool isSucceed = this.SutControlAdapter.UnholdFiles(siteUrl, holdName, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Un-hold all the items of the list library {0} from the hold {1} in the site{2}. Actual result is {3}", documentLibary, holdName, siteUrl, isSucceed ? "success" : "failure"));

                isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actually result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));
            }
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is a storage location configured for routing content when the required properties are not specified.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC19_SubmitFile_InvalidArgument()
        {
            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            // Convert the random file content into Base64Encoded output.
            byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

            // Call SubmitFile on a repository that is a storage location configured for routing content when the required properties are not specified, expect the server responses InvalidArgument.
            SubmitFileResult submitFileResult = this.Adapter.SubmitFile(fileToSubmit, null, this.DocumentContentTypeName, this.GetOriginalSaveLocation(this.DocumentLibraryUrlOfAppendUniqueSuffix), this.SubmitUserName);

            // Add the log information.
            Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "The ResultCode is {0}", submitFileResult.ResultCode.ToString());

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1991
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.InvalidArgument,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     1991,
                     @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] If required parameters were not specified, the protocol server MUST set the ResultCode to InvalidArgument and return.");

            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.InvalidArgument,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     251,
                     @"[In SubmitFileResultCode] [If] The operation [SubmitFile] failed because of an invalid argument. [SubmitFileResultCode] Value is InvalidArgument.");
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is a storage location configured for routing content and all the parameters are valid.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC20_SubmitFile_Success_CommonProperties()
        {
            if (!Common.Common.IsRequirementEnabled(1049, this.Site)
                && !Common.Common.IsRequirementEnabled(215001, this.Site)
                && !Common.Common.IsRequirementEnabled(1050, this.Site))
            {
                this.Site.Assume.Inconclusive("In the test case MSOFFICIALFILE_S01_TC20_SubmitFile_Success_CommonProperties, at least one of the requirements MS-OFFCIIALFILE_R1049, MS-OFFICIALFILE_R215001, MS-OFFICIALFILE_R1050 needs to be enable.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;
            this.Adapter.IntializeService(paras);

            try
            {
                // Properties of the file to submit which will contain all properties.
                RecordsRepositoryProperty[] fileProperties = this.ConstructAllProperties();

                // Convert the binary input into Base64 UnEncoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

                // Call SubmitFile on a repository that is a storage location configured for routing content and with all required and common properties, expect the server responds Success and non-empty result URL.
                SubmitFileResult submitFileResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, this.GenerateRandomTextFileName(), this.SubmitUserName);

                // Verify the success related requirements for SubmitFile.
                this.VerifySubmitFileSucceed(submitFileResult);
                if (Common.Common.IsRequirementEnabled(1049, this.Site))
                {
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1049
                    Site.CaptureRequirementIfIsNull(
                             submitFileResult.CustomProcessingResult,
                             "MS-OFFICIALFILE",
                             1049,
                             @"[In Appendix B: Product Behavior][The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:  If the properties element contains all of the following properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery, _dlc_hold_searchcontexturl] Implementation does not include the CustomProcessingResult element. <17> Section 3.1.4.6:  Office SharePoint Server 2007 does not include the CustomProcessingResult element.");
                }

                if (Common.Common.IsRequirementEnabled(215001, this.Site))
                {
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R215001
                    Site.Assert.IsNotNull(
                      submitFileResult.CustomProcessingResult,
                      "For Requirement R215001, theCustomProcessingResult should be returned");

                    Site.CaptureRequirementIfAreEqual(
                             HoldProcessingResult.Success,
                             submitFileResult.CustomProcessingResult.HoldProcessingResult,
                             "MS-OFFICIALFILE",
                             215001,
                             @"[In Appendix B: Product Behavior][The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:  If the properties element contains all of the following properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery, _dlc_hold_searchcontexturl] Implementation does include CustomProcessingResult element and set the CustomProcessingResult.HoldsProcessingResult element to Success. (Microsoft SharePoint Server 2010 and above follow this behavior.)");

                    Site.CaptureRequirementIfAreEqual(
                             HoldProcessingResult.Success,
                             submitFileResult.CustomProcessingResult.HoldProcessingResult,
                             "MS-OFFICIALFILE",
                             258,
                             @"[In HoldProcessingResult] [If] The processing of a legal hold is successful. [HoldProcessingResult] Value is Success.");
                }

                if(Common.Common.IsRequirementEnabled(360,this.Site))
                {
                    Site.CaptureRequirementIfAreEqual(
                         HoldProcessingResult.Success,
                         submitFileResult.CustomProcessingResult.HoldProcessingResult,
                         "MS-OFFICIALFILE",
                         360,
                         @"[In Appendix B: Product Behavior] Implementation does include CustomProcessingResult element and set the CustomProcessingResult.HoldsProcessingResult. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                // Remove one of the common properties with the property name _dlc_hold_url
                fileProperties = this.ConstructAllProperties();
                fileProperties = fileProperties.Where(p => p.Name != "_dlc_hold_url").ToArray();

                // Call SubmitFile on a repository that is a storage location configured for routing content and not all the common properties which will miss _dlc_hold_url property.
                submitFileResult = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.DocumentContentTypeName, this.GenerateRandomTextFileName(), this.SubmitUserName);
                if (Common.Common.IsRequirementEnabled(1050, this.Site))
                {
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1050
                    Site.CaptureRequirementIfIsNull(
                             submitFileResult.CustomProcessingResult,
                             "MS-OFFICIALFILE",
                             1050,
                             @"[In Appendix B: Product Behavior][The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows: If, however, at least one but not all of the properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery, _dlc_hold_searchcontexturl are contained in the properties element] Implementation does not include the CustomProcessingResult element.");
                }
            }
            finally
            {
                string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DocumentLibraryName", this.Site);
                string holdName = Common.Common.GetConfigurationPropertyValue("HoldName", this.Site);

                bool isSucceed = this.SutControlAdapter.UnholdFiles(siteUrl, holdName, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Un-hold all the items of the list library {0} from the hold {1} in the site{2}. Actual result is {3}", documentLibary, holdName, siteUrl, isSucceed ? "success" : "failure"));

                isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actually result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));
            }
        }

        /// <summary>
        /// This test case is used to test SubmitFile on a repository that is configured for routing content and no configured routing rule for the specified content type name parameter.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S01_TC21_SubmitFile_NoApplicableRule()
        {
            if (!Common.Common.IsRequirementEnabled(203001, this.Site)
                && !Common.Common.IsRequirementEnabled(204001, this.Site))
            {
                this.Site.Assume.Inconclusive("In the test case MSOFFICIALFILE_S01_TC21_SubmitFile_NoApplicableRule, at least one of the requirements MS-OFFCIIALFILE_R203001 and MS-OFFICIALFILE_R204001 needs to be enable.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            try
            {
                // Properties of the file to submit which will contain all required properties.
                RecordsRepositoryProperty[] fileProperties = this.ConstructAllRequiredProperties();

                // Convert the random file content into Base64Encoded output.
                byte[] fileToSubmit = AdapterHelper.EncodeToBase64(Common.Common.GenerateResourceName(Site, "FileContent"));

                // Call SubmitFile on a repository that is configured for routing content and no configured routing rule for the specified content type name parameter, expect the server responses MoreInformation.
                SubmitFileResult submitFileResult1 = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.NotSupportedContentTypeName, this.GenerateRandomTextFileName(), this.SubmitUserName);

                if (Common.Common.IsRequirementEnabled(203001, this.Site))
                {
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R203001
                    Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                             SubmitFileResultCode.MoreInformation,
                             submitFileResult1.ResultCode,
                             "MS-OFFICIALFILE",
                             203001,
                             @"[In Appendix B: Product Behavior] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] Implementation does set the ResultCode element to MoreInformation when no applicable rule is found. (Microsoft SharePoint Server 2010 and above follow this behavior.)");

                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R249
                    Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                             SubmitFileResultCode.MoreInformation,
                             submitFileResult1.ResultCode,
                             "MS-OFFICIALFILE",
                             249,
                             @"[In SubmitFileResultCode] [If] The operation [SubmitFile] is successful but further action is needed. [SubmitFileResultCode] Value is MoreInformation.");
                }

                if (Common.Common.IsRequirementEnabled(204001, this.Site))
                {
                    Site.CaptureRequirementIfIsNotNull(
                             submitFileResult1.ResultUrl,
                             "MS-OFFICIALFILE",
                             204001,
                             @"[In Appendix B: Product Behavior] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] Implementation does set the ResultUrl element to an implementation-specific URL to enter more information about the submission when no applicable rule is found. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                // Call SubmitFile on a repository that is configured for routing content and no configured routing rule for the specified content type name parameter 
                // with all of  the common properties _dlc_hold_url, _dlc_hold_comments, _dlc_hold_id, _dlc_hold_searchqquery and _dlc_hold_searchcontexturl, expect the server responses MoreInformation.
                fileProperties = this.ConstructAllProperties();
                SubmitFileResult submitFileResult2 = this.Adapter.SubmitFile(fileToSubmit, fileProperties, this.NotSupportedContentTypeName, this.GenerateRandomTextFileName(), this.SubmitUserName);

                if (Common.Common.IsRequirementEnabled(203001, this.Site))
                {
                    // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R203001
                    Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                             SubmitFileResultCode.MoreInformation,
                             submitFileResult2.ResultCode,
                             "MS-OFFICIALFILE",
                             203001,
                             @"[In Appendix B: Product Behavior] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] Implementation does set the ResultCode element to MoreInformation when no applicable rule is found. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                if (Common.Common.IsRequirementEnabled(204001, this.Site))
                {
                    Site.CaptureRequirementIfIsNotNull(
                             submitFileResult2.ResultUrl,
                             "MS-OFFICIALFILE",
                             204001,
                             @"[In Appendix B: Product Behavior] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:] Implementation does set the ResultUrl element to an implementation-specific URL to enter more information about the submission when no applicable rule is found. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }
            }
            finally
            {
                string siteUrl = Common.Common.GetConfigurationPropertyValue("EnableContentOrganizerRecordsCenterSite", this.Site);
                string documentLibary = Common.Common.GetConfigurationPropertyValue("DefaultLibraryName", this.Site);
                string holdName = Common.Common.GetConfigurationPropertyValue("HoldName", this.Site);

                bool isSucceed = this.SutControlAdapter.UnholdFiles(siteUrl, holdName, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Un-hold all the items of the list library {0} from the hold {1} in the site{2}. Actual result is {3}", documentLibary, holdName, siteUrl, isSucceed ? "success" : "failure"));

                isSucceed = this.SutControlAdapter.DeleteAllFiles(siteUrl, documentLibary);
                this.Site.Assert.IsTrue(
                    isSucceed,
                    string.Format("Delete the submitted files in the site {0} and library {1} should succeed. Actually result is {2}", siteUrl, documentLibary, isSucceed ? "success" : "failure"));
            }
        }

        /// <summary>
        /// This method is used to verify all the related requirements when the submit file succeeds.
        /// </summary>
        /// <param name="submitFileResult">The SubmitFileResult respond by the protocol server.</param>
        private void VerifySubmitFileSucceed(SubmitFileResult submitFileResult)
        {
            // Add the log information.
            Site.Log.Add(
                Microsoft.Protocols.TestTools.LogEntryKind.Comment,
                "ResultUrl is {0}, ResultCode is {1}.",
                submitFileResult.ResultUrl,
                submitFileResult.ResultCode);

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R213
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.Success,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     213,
                     @"[In SubmitFile] [The protocol client sends a SubmitFileSoapIn request WSDL message, and the protocol server MUST respond with a SubmitFileSoapOut response WSDL message, as follows:]
[If implementation-specific errors occur while determining the storage location for the submission or while storing the file, the protocol server MUST set the ResultCode element to UnknownError and return.] Otherwise, , the protocol server MUST set the ResultCode element to Success [and SHOULD<16> set the ResultUrl element to a non-empty HTML encoded URL to retrieve the stored file.]");

            // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R248
            Site.CaptureRequirementIfAreEqual<SubmitFileResultCode?>(
                     SubmitFileResultCode.Success,
                     submitFileResult.ResultCode,
                     "MS-OFFICIALFILE",
                     248,
                     @"[In SubmitFileResultCode] [If] The operation [SubmitFile] is successful. [SubmitFileResultCode] Value is Success.");

            if (Common.Common.IsRequirementEnabled(1067, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                         submitFileResult.ResultUrl,
                         "MS-OFFICIALFILE",
                         1067,
                         @"[In Appendix B: Product Behavior] Implementation does set the ResultCode element to Success and set the ResultUrl element to a non-empty HTML encoded URL to retrieve the stored file if implementation-specific errors doesn't occur while determining the storage location for the submission or while storing the file. (Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }

            if (Common.Common.IsRequirementEnabled(1066, this.Site))
            {
                // Verify MS-OFFICIALFILE requirement: MS-OFFICIALFILE_R1066
                Site.CaptureRequirementIfIsNotNull(
                         submitFileResult.AdditionalInformation,
                         "MS-OFFICIALFILE",
                         1066,
                         @"[In Appendix B: Product Behavior] Implementation does set the ResultUrl element to a non-empty HTML encoded URL in the AdditionalInformation element to retrieve the stored file. <16> Section 3.1.4.6:  Office SharePoint Server 2007 returns the URL in the AdditionalInformation element.");
            }
        }
        #endregion
    }
}