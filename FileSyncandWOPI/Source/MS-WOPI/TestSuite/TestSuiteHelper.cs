namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.IO;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to support required functions for test class initialization and clean up
    /// </summary>
    public class TestSuiteHelper : HelperBase
    {
        #region Variables

        /// <summary>
        /// A string value represents the protocol short name for the MS-WOPI.
        /// </summary>
        private const string WopiProtocolShortName = "MS-WOPI";

        /// <summary>
        /// A string value represents the protocol short name for the shared test cases, it is used in runtime. If plan to run the shared test cases, the WOPI server must implement the MS-FSSHTTP.
        /// </summary>
        private const string SharedTestCasesProtocolShortName = "MS-FSSHTTP-FSSHTTPB";

        /// <summary>
        /// Represents a IMS_WOPISUTManageCodeControlAdapter type instance.
        /// </summary>
        private static IMS_WOPIManagedCodeSUTControlAdapter wopiSutManagedCodeControlAdapter;

        /// <summary>
        /// Represents the IMS_WOPISUTControlAdapter instance.
        /// </summary>
        private static IMS_WOPISUTControlAdapter wopiSutControlAdapter;

        /// <summary>
        /// Represents the IMS_WOPIAdapter instance.
        /// </summary>
        private static IMS_WOPIAdapter wopiProtocolAdapter;

        /// <summary>
        /// A string value represents the current test client name which the test suite run on.
        /// </summary>
        private static string currentTestClientName;

        /// <summary>
        /// A string value represents the progId which is used in discovery process in order to make the WOPI server enable folder level visit ability when it receive the progId from the discovery response. 
        /// </summary>
        private static string progId = string.Empty;

        /// <summary>
        /// A string value represents the URL of the relative source file.
        /// </summary>
        private static string relativeSourceFileUrl = string.Empty;

        /// <summary>
        /// A bool value represents whether the ShareCaseHelper has been initialized. The value 'true' means it has been initialized.
        /// </summary>
        private static bool hasInitializedHelperStatus = false;

        /// <summary>
        /// A bool value represents whether the test suite merge the configuration files.
        /// </summary>
        private static bool hasMergeWOPIPtfConfigFile = false;

        /// <summary>
        /// A long type value represents the clean up counter for the discovery process. In each trigger WOPI discovery request purpose process, this counter will increment, and decrement in each discovery clean up purpose process.
        /// </summary>
        private static long cleanUpDiscoveryStatusCounter = 0;

        /// <summary>
        /// A string value represents the current WCF endpoint name when running the shared test cases. 
        /// </summary>
        private static string currentSharedTestCasesEndpointName;

        /// <summary>
        /// A bool value represents whether the current SUT support the MS-WOPI protocol.
        /// </summary>
        private static bool isCurrentSUTSupportWOPI = false;

        /// <summary>
        /// A bool value represents whether the test suite has checked that the current SUT support the MS-WOPI protocol.
        /// </summary>
        private static bool hasCheckSupportWOPI = false;

        #endregion

        /// <summary>
        /// Prevents a default instance of the TestSuiteHelper class from being created
        /// </summary>
        private TestSuiteHelper()
        {
        }

        /// <summary>
        /// Gets a value indicating whether the ShareCaseHelper has been initialized. The value 'true' means it has been initialized.
        /// </summary>
        public static bool HasInitialized
        {
            get
            {
                return hasInitializedHelperStatus;
            }
        }

        /// <summary>
        /// Gets the IMS_WOPISUTControlAdapter instance.
        /// </summary>
        public static IMS_WOPISUTControlAdapter WOPISutControladapter
        {
            get
            {
                if (null == wopiSutControlAdapter)
                {
                    throw new InvalidOperationException("Should call the [InitializeHelper] method to initial this helper.");
                }

                return wopiSutControlAdapter;
            }
        }

        /// <summary>
        /// Gets the IMS_WOPIManagedCodeSUTControlAdapter instance.
        /// </summary>
        public static IMS_WOPIManagedCodeSUTControlAdapter WOPIManagedCodeSUTControlAdapter
        {
            get
            {
                if (null == wopiSutManagedCodeControlAdapter)
                {
                    throw new InvalidOperationException("Should call the [InitializeHelper] method to initial this helper.");
                }

                return wopiSutManagedCodeControlAdapter;
            }
        }

        /// <summary>
        /// Gets the IMS_WOPIAdapter instance.
        /// </summary>
        public static IMS_WOPIAdapter WOPIProtocolAdapter
        {
            get
            {
                if (null == wopiProtocolAdapter)
                {
                    throw new InvalidOperationException("Should call the [InitializeHelper] method to initial this helper.");
                }

                return wopiProtocolAdapter;
            }
        }

        /// <summary>
        /// Gets the name of current test client.
        /// </summary>
        public static string CurrentTestClientName
        {
            get
            {
                if (string.IsNullOrEmpty(currentTestClientName))
                {
                    throw new InvalidOperationException("Should call the [InitializeHelper] method to initial this helper.");
                }

                return currentTestClientName;
            }
        }

        /// <summary>
        /// This method is used to initialize the share test case helper. This method will also initialize all helpers which are required to initialize during test suite running.
        /// </summary>
        /// <param name="siteInstance">A parameter represents the ITestSite instance.</param>
        public static void InitializeHelper(ITestSite siteInstance)
        {
            TestSuiteHelper.CheckInputParameterNullOrEmpty<ITestSite>(siteInstance, "siteInstance", "InitializeHelper");

            if (string.IsNullOrEmpty(currentTestClientName))
            {
                currentTestClientName = Common.GetConfigurationPropertyValue("TestClientName", siteInstance);
            }

            if (null == wopiSutControlAdapter)
            {
                wopiSutControlAdapter = siteInstance.GetAdapter<IMS_WOPISUTControlAdapter>();
            }

            if (null == wopiSutManagedCodeControlAdapter)
            {
                wopiSutManagedCodeControlAdapter = siteInstance.GetAdapter<IMS_WOPIManagedCodeSUTControlAdapter>();
            }

            if (null == wopiProtocolAdapter)
            {
                wopiProtocolAdapter = siteInstance.GetAdapter<IMS_WOPIAdapter>();
            }

            InitializeRequiredHelpers(wopiSutManagedCodeControlAdapter, siteInstance);

            if (string.IsNullOrEmpty(relativeSourceFileUrl))
            {
                relativeSourceFileUrl = Common.GetConfigurationPropertyValue("NormalFile", siteInstance);
            }

            progId = Common.GetConfigurationPropertyValue("ProgIdForDiscoveryProcess", siteInstance);

            // Setting the endpoint name according to the current http transport.
            if (string.IsNullOrEmpty(currentSharedTestCasesEndpointName))
            {
                TransportProtocol currentTransport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", siteInstance);
                switch (currentTransport)
                {
                    case TransportProtocol.HTTP:
                        {
                            currentSharedTestCasesEndpointName = Common.GetConfigurationPropertyValue("SharedTestCaseEndPointNameForHTTP", siteInstance);
                            break;
                        }

                    case TransportProtocol.HTTPS:
                        {
                            currentSharedTestCasesEndpointName = Common.GetConfigurationPropertyValue("SharedTestCaseEndPointNameForHTTPS", siteInstance);
                            break;
                        }

                    default:
                        {
                            throw new InvalidOperationException(string.Format("The test suite only support HTTP or HTTPS transport. Current:[{0}]", currentTransport));
                        }
                }
            }

            // Set the protocol name of current test suite
            siteInstance.DefaultProtocolDocShortName = WopiProtocolShortName;

            hasInitializedHelperStatus = true;
        }

        /// <summary>
        /// This method is used to initialize the shared test cases' context.
        /// </summary>
        /// <param name="requestFileUrl">A parameter represents the file URL.</param>
        /// <param name="userName">A parameter represents the user name we used.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain.</param>
        /// <param name="celloperationType">A parameter represents the type of CellStore operation which is used to determine different initialize logic.</param>
        /// <param name="site">A parameter represents the site.</param>
        public static void InitializeContextForShare(string requestFileUrl, string userName, string password, string domain, CellStoreOperationType celloperationType, ITestSite site)
        {
            SharedContext context = SharedContext.Current;
            string normalTargetUrl = string.Empty;

            switch (celloperationType)
            {
                case CellStoreOperationType.NormalCellStore:
                    {
                        normalTargetUrl = requestFileUrl;
                        context.OperationType = OperationType.WOPICellStorageRequest;
                        break;
                    }

                case CellStoreOperationType.RelativeAdd:
                    {
                        // For relative adding a file, the WOPI resource file URL should be an existed file, and the file should be added in a location where the relative source file is located. 
                        normalTargetUrl = relativeSourceFileUrl;
                        context.OperationType = OperationType.WOPICellStorageRelativeRequest;
                        break;
                    }

                case CellStoreOperationType.RealativeModified:
                    {
                        normalTargetUrl = requestFileUrl;
                        context.OperationType = OperationType.WOPICellStorageRelativeRequest;
                        break;
                    }
            }

            // Convert the request file url to WOPI format URL so that the MS-FSSHTTP test cases will use WOPI format URL to send request.
            string wopiTargetUrl = wopiSutManagedCodeControlAdapter.GetWOPIRootResourceUrl(normalTargetUrl, WOPIRootResourceUrlType.FileLevel, userName, password, domain);

            context.TargetUrl = wopiTargetUrl;
            context.IsMsFsshttpRequirementsCaptured = false;
            context.Site = site;

            // Only set the X-WOPI-RelativeTarget header value in relative CellStorage operation.
            if (CellStoreOperationType.NormalCellStore != celloperationType)
            {
                // Read the file name from the target resource URL.
                context.XWOPIRelativeTarget = GetFileNameFromFullUrl(requestFileUrl);
            }

            // Generate common headers
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetUrl);
            context.XWOPIProof = commonHeaders["X-WOPI-Proof"];
            context.XWOPIProofOld = commonHeaders["X-WOPI-ProofOld"];
            context.XWOPITimeStamp = commonHeaders["X-WOPI-TimeStamp"];
            context.XWOPIAuthorization = commonHeaders["Authorization"];
            context.EndpointConfigurationName = currentSharedTestCasesEndpointName;
            context.UserName = userName;
            context.Password = password;
            context.Domain = domain;
        }

        /// <summary>
        /// This method is used to merge the configurations which are used by shared test cases.
        /// </summary>
        /// <param name="siteInstance">A parameter represents the site.</param>
        public static void MergeConfigurationFileForShare(ITestSite siteInstance)
        {
            siteInstance.DefaultProtocolDocShortName = SharedTestCasesProtocolShortName;

            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", siteInstance);

            // Merge the common configuration.
            Common.MergeGlobalConfig(commonConfigFileName, siteInstance);

            // Merge the should may configuration file.
            Common.MergeSHOULDMAYConfig(siteInstance);

            // Roll back the short name in the ITestSite instance.
            siteInstance.DefaultProtocolDocShortName = WopiProtocolShortName;
        }

        /// <summary>
        /// A method is used to perform the discovery process
        /// </summary>
        /// <param name="currentTestClient">A parameter represents the current test client which is listening the discovery request.</param>
        /// <param name="sutControllerAdapterInstance">A parameter represents the IMS_WOPISUTControlAdapter instance which is used to make the WOPI server perform sending discovery request to the discovery listener.</param>
        /// <param name="siteInstance">A parameter represents the ITestSite instance which is used to get the test context.</param>
        public static void PerformDiscoveryProcess(string currentTestClient, IMS_WOPISUTControlAdapter sutControllerAdapterInstance, ITestSite siteInstance)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(currentTestClient, "currentTestClient", "PerformDiscoveryProcess");
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<IMS_WOPISUTControlAdapter>(sutControllerAdapterInstance, "sutControllerAdapterInstance", "PerformDiscoveryProcess");
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<ITestSite>(siteInstance, "siteInstance", "PerformDiscoveryProcess");

            // If the test class invoke this, means the test class will uses the WOPI discovery binding. The test suite will count all WOPI discovery usage of test classes
            System.Threading.Interlocked.Increment(ref cleanUpDiscoveryStatusCounter);

            // Start the listener, if the listen thread has been start, the DiscoverProcessHelper will not start any new listen thread.
            DiscoveryProcessHelper.StartDiscoveryListen(currentTestClient, progId);

            // Initialize the WOPI Discovery process so that the WOPI server will use the test suite as WOPI client.
            if (!DiscoveryProcessHelper.HasPerformDiscoveryProcessSucceed)
            {
                DiscoveryProcessHelper.PerformDiscoveryProcess(currentTestClient, sutControllerAdapterInstance);
            }
        }

        /// <summary>
        /// A method is used to clean the WOPI discovery process for the WOPI server.
        /// </summary>
        /// <param name="currentTestClient">A parameter represents the test client name which is running the test suite. This test client act as WOPI client and response discovery request from the WOPI server successfully.</param>
        /// <param name="sutControllerAdapterInstance">A parameter represents the IMS_WOPISUTControlAdapter instance which is used to make the WOPI server perform sending discovery request to the discovery listener.</param>
        public static void CleanUpDiscoveryProcess(string currentTestClient, IMS_WOPISUTControlAdapter sutControllerAdapterInstance)
        {
            // If the current SUT does not support MS-WOPI, the test suite will not perform any discovery process logics.
            if (!isCurrentSUTSupportWOPI)
            {
                return;
            }

            // If the test class invoke this, means the test class will uses the binding.
            long currentCleanUpStatusCounter = System.Threading.Interlocked.Decrement(ref cleanUpDiscoveryStatusCounter);

            if (currentCleanUpStatusCounter > 0)
            {
                return;
            }
            else if (0 == currentCleanUpStatusCounter)
            {
                // Clean up the WOPI discovery process record from so the WOPI server. The test suite act as WOPI client.
                if (DiscoveryProcessHelper.NeedToCleanUpDiscoveryRecord)
                {
                    DiscoveryProcessHelper.CleanUpDiscoveryRecord(currentTestClient, sutControllerAdapterInstance);
                }

                // Dispose the discovery request listener.
                DiscoveryProcessHelper.DisposeDiscoveryListener();
            }
            else
            {
                throw new InvalidOperationException(string.Format("The discovery clean up counter should not be less than zero. current value[{0}]", currentCleanUpStatusCounter));
            }
        }

        /// <summary>
        /// This method is used to get the file name from the URL.
        /// </summary>
        /// <param name="fullUrlOfFile">A parameter represents the full URL of the file.</param>
        /// <returns>A parameter represents the file name.</returns>
        public static string GetFileNameFromFullUrl(string fullUrlOfFile)
        {
            // Ensure the file name exists.
            if (string.IsNullOrEmpty(fullUrlOfFile))
            {
                throw new ArgumentNullException("fullUrlOfFile");
            }

            // Get the file name.
            Uri currentFilePath;
            if (!Uri.TryCreate(fullUrlOfFile, UriKind.Absolute, out currentFilePath))
            {
                throw new UriFormatException("The [fullUrlOfFile] parameter must be a valid absolute URL.");
            }

            string fileName = Path.GetFileName(currentFilePath.LocalPath);
            if (string.IsNullOrEmpty(fileName))
            {
                throw new InvalidOperationException(string.Format("Could not get the file name from the file path:[{0}].", currentFilePath.OriginalString));
            }

            return fileName;
        }

        /// <summary>
        /// This method is used to merge the configuration of the WOPI PTF configuration files.
        /// </summary>
        /// <param name="siteInstance">A parameter represents the site.</param>
        public static void MergeWOPIPtfConfigFiles(ITestSite siteInstance)
        {
            if (!hasMergeWOPIPtfConfigFile)
            {
                // Set the default short name for WOPI.
                siteInstance.DefaultProtocolDocShortName = WopiProtocolShortName;

                // Merge the common configuration.
                string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", siteInstance);
                Common.MergeGlobalConfig(conmmonConfigFileName, siteInstance);

                // Merge the SHOULDMAY configuration
                Common.MergeSHOULDMAYConfig(siteInstance);
                hasMergeWOPIPtfConfigFile = true;
            }
        }

        /// <summary>
        /// A method is used to check whether the SUT product supports the MS-WOPI protocol. If the SUT does not support, this method will raise an inconclusive assertion.
        /// </summary>
        /// <param name="site">A parameter represents the site.</param>
        public static void PerformSupportProductsCheck(ITestSite site)
        {
            TestSuiteHelper.CheckInputParameterNullOrEmpty<ITestSite>(site, "siteInstance", "PerformSupportProductsCheck");
            if (!hasCheckSupportWOPI)
            {
                isCurrentSUTSupportWOPI = Common.GetConfigurationPropertyValue<bool>("MS-WOPI_Supported", site);
                hasCheckSupportWOPI = true;
            }

            if (!isCurrentSUTSupportWOPI)
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", site);
                site.Assume.Inconclusive(@"The server does not support this specification [MS-WOPI]. It is determined by ""MS-WOPI_Supported"" SHOULDMAY property of the [{0}_{1}_SHOULDMAY.deployment.ptfconfig] configuration file.", WopiProtocolShortName, currentSutVersion);
            }
        }

        /// <summary>
        /// A method is used to check whether the WOPI server Cobalt feature, this feature depends on that the WOPI server whether implements MS-FSSHTTP. If current WOPI server does support Cobalt feature, this method will capture related requirement R961, otherwise this method will raise an inconclusive assertion. 
        /// </summary>
        /// <param name="siteInstance">A parameter represents the site instance.</param>
        public static void PerformSupportCobaltCheck(ITestSite siteInstance)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<ITestSite>(siteInstance, "siteInstance", "PerformSupportCobaltCheck");
            if (!Common.IsRequirementEnabled("MS-WOPI", 961, siteInstance))
            {
                siteInstance.Assert.Inconclusive(@"The WOPI server does not support the Cobalt feature. To supporting this feature, the WOPI server must implement the MS-FSSHTTP protocol.It is determined by ""R961Enabled_MS-WOPI"" SHOULDMAY property.");
            }
        }

        /// <summary>
        /// A method used to initialize the test suite. It is used in test class level initialization.
        /// </summary>
        /// <param name="testSite">A parameter represents the ITestSite instance which contain the test context information.</param>
        public static void InitializeTestSuite(ITestSite testSite)
        {
            if (null == testSite)
            {
                throw new ArgumentNullException("testSite");
            }

            try
            {
                if (!TestSuiteHelper.HasInitialized)
                {
                    TestSuiteHelper.InitializeHelper(testSite);
                }

                TestSuiteHelper.PerformDiscoveryProcess(
                                                        currentTestClientName,
                                                        wopiSutControlAdapter,
                                                        testSite);
            }
            catch (Exception)
            {
                TestSuiteHelper.CleanUpDiscoveryProcess(currentTestClientName, wopiSutControlAdapter);
                throw;
            }
        }

        /// <summary>
        /// A method is used to verify whether test suite run in support products. If the value of "SutVersion" property in common configuration file is not included in "SupportProducts" property in "MS-WOPI_TestSuite.deployment.ptfconfig", the "MS-WOPI_Supported" SHOULDMAY property will always equal to false. And the unsupported initialization logic will not be executed in unsupported products.
        /// </summary>
        /// <param name="siteInstance">A parameter represents the site instance.</param>
        /// <returns>Return 'true' indicating the test suite is running in support products. The initialization logics should be performed.</returns>
        public static bool VerifyRunInSupportProducts(ITestSite siteInstance)
        {
            if (null == siteInstance)
            {
                throw new ArgumentNullException("siteInstance");
            }

            TestSuiteHelper.MergeWOPIPtfConfigFiles(siteInstance);
            return Common.GetConfigurationPropertyValue<bool>("MS-WOPI_Supported", siteInstance);
        }

        /// <summary>
        /// A method is used to initialize all required helpers when calling the SharedCase helper. Normally, all this method will try to initial all helpers which are required to initialize.  
        /// </summary>
        ///  <param name="sutManagedCodeControlAdpater">A parameter represents the IMS_WOPISUTManageCodeControlAdapter instance which is used to convert the normal resource URL to WOPI format URL.</param>
        /// <param name="siteInstance">A parameter represents the ITestSite instance.</param>
        private static void InitializeRequiredHelpers(IMS_WOPIManagedCodeSUTControlAdapter sutManagedCodeControlAdpater, ITestSite siteInstance)
        {
            if (!TokenAndRequestUrlHelper.HasInitialized)
            {
                TokenAndRequestUrlHelper.InitializeHelper(sutManagedCodeControlAdpater, siteInstance);
            }
        }
    }
}