namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Threading;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class contains test case definitions related to ExportWeb operation.
    /// </summary>
    [TestClass]
    public class S01_MigrateSite : TestClassBase
    {
        /// <summary>
        /// An instance of protocol adapter class.
        /// </summary>
        private IMS_SITESSAdapter sitessAdapter;

        /// <summary>
        /// An instance of SUT control adapter class.
        /// </summary>
        private IMS_SITESSSUTControlAdapter sutAdapter;

        /// <summary>
        /// The name of the sub site to be imported.
        /// </summary>
        private string newSubsite = string.Empty;

        /// <summary>
        /// A boolean value indicates whether a web is imported successfully by the ImportWeb operation.
        /// </summary>
        private bool isWebImportedSuccessfully = false;

        #region Test Suite Initialization & Cleanup

        /// <summary>
        /// Test Suite Initialization.
        /// </summary>
        /// <param name="testContext">The test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Test Suite Cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        #region Scenario 1 Migrate a site

        /// <summary>
        /// This test case is designed to verify the ExportWeb and ImportWeb operations when migrating a site successfully.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC01_MigratingSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = siteUrl + "/" + this.newSubsite;
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string logPath = string.Empty;
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string importJobName = Constants.ImportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { dataPath + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            int exportWebResult = 0;
            int importWebResult = 0;
            string[] exportWebFiles = null;
            string[] expectedWebFiles = null;
            string exportWebStatusCode = string.Empty;
            string importWebStatusCode = string.Empty;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid parameters, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Format the expected file names in the document library. Only one cmp file (i.e. ExportWeb.cmp) is expected to be exported in the document library.
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.CmpExtension,
                exportJobName + Constants.SntExtension
            };

            exportWebStatusCode = this.sutAdapter.GetStatusCode(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site), exportJobName + Constants.SntExtension);

            #region Capture requirements

            if (Common.IsRequirementEnabled(329, this.Site))
            {
                // If a status code is included in the file which is generated by ExportWeb, R329 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R329");

                // Verify MS-SITESS requirement: MS-SITESS_R329
                Site.CaptureRequirementIfAreNotEqual<string>(
                    string.Empty,
                    exportWebStatusCode,
                    329,
                    @"[In Appendix B: Product Behavior][<3> Section 3.1.4.2:] The file [which contains the result of the export operation] includes a status code indicating the success of the operation or an error code. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(328, this.Site))
            {
                // If a status code is included in the file which is generated by ExportWeb, it means the file 
                // that contains the result of the export operation has been successfully generated. R328 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R328");

                // Verify MS-SITESS requirement: MS-SITESS_R328
                Site.CaptureRequirementIfAreNotEqual<string>(
                    string.Empty,
                    exportWebStatusCode,
                    328,
                    @"[In Appendix B: Product Behavior] <3> Section 3.1.4.2: Implementation does creates a file that contains the result of the export operation in the server location specified in the request message. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(433001, this.Site))
            {
                for (int i = 0; i < expectedWebFiles.Length; i++)
                {
                    // If a status code is included in the file which is generated by ExportWeb, it means the file 
                    // that contains the result of the export operation has been successfully generated. R433001 can be captured.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R433001");

                    // Verify MS-SITESS requirement: MS-SITESS_R433001
                    Site.CaptureRequirementIfAreNotEqual<string>(
                        string.Empty,
                        exportWebFiles[i],
                        433001,
                        @"[In Appendix B: Product Behavior] [<3> Section 3.1.4.2:] If the export operation succeeds, this file also includes the list of the content migration package files created. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
                }
            }

            // Verify whether the exported file is Content Migration Package(.cmp) file. Since all the exported file and log file are compared, the file name of the exported file is also verified.
            bool isCmpFile = AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles);

            // If export web files is .cmp file, R71 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R71, the expected Content Migration Package(.cmp) file is {0}, and the exported file name list is {1}", expectedWebFiles[0], files);

            // Verify MS-SITESS requirement: MS-SITESS_R71
            Site.CaptureRequirementIfIsTrue(
                isCmpFile,
                71,
                @"[In ExportWeb] [jobName:] The server MUST append the file extension .cmp to the jobName to form the file name for the first content migration package file.");

            // If export web files is .cmp file, R60 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R60, the expected Content Migration Package(.cmp) file is {0}", expectedWebFiles[0]);

            // Verify MS-SITESS requirement: MS-SITESS_R60
            Site.CaptureRequirementIfIsTrue(
                isCmpFile,
                60,
                @"[In ExportWeb] Upon the successful completion of the export operation, the Content Migration Package file(s) MUST be created in the server location specified in the request message.");

            #endregion Capture requirements

            logPath = dataPath;

            // Invoke the ImportWeb operation with valid parameters, 1 is expected to be returned.
            importWebResult = this.sitessAdapter.ImportWeb(importJobName, importUrl, dataFiles, logPath, true, true);

            #region Capture requirements

            // If 1 is returned, requirements related with ImportWeb pending can be verified.
            this.VerifyImportWebInProgress(importWebResult);
            this.isWebImportedSuccessfully = true;

            DateTime beforeLoop = DateTime.Now;
            TimeSpan waitTime = new TimeSpan();
            int repeatTime = 0;
            int totalRepeatTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ImportWebRepeatTime, this.Site));
            do
            {
                // It is assumed that the server will generate the status code of ImportWeb operation after the preconfigured time period.
                int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ImportWebWaitTime, this.Site));
                Thread.Sleep(1000 * sleepTime);

                // Get the status code in the log file.
                importWebStatusCode = this.sutAdapter.GetStatusCode(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site), importJobName + Constants.SntExtension);

                // If the returned status code is not null, the operation is considered as succeed.
                if (!string.IsNullOrEmpty(importWebStatusCode))
                {
                    break;
                }

                // If the server could not generate the status code after the time period ImportWebWaitTime for some unknown reasons, for example, limit of server resources, try to repeat this sequence again.
                repeatTime++;
            }
            while (repeatTime < totalRepeatTime);
            waitTime = DateTime.Now - beforeLoop;

            // If the server still does not generate the status code after repeating, the operation is considered as failed.
            if (string.IsNullOrEmpty(importWebStatusCode))
            {
                Site.Assert.Fail("The result of the import operation was not received after {0} seconds", waitTime);
            }

            if (Common.IsRequirementEnabled(340, this.Site))
            {
                // If a status code is included in the file which is generated by ImportWeb, R340 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R340");

                // Verify MS-SITESS requirement: MS-SITESS_R340
                // Since the returned status code is not empty, this requirement can be captured directly.
                Site.CaptureRequirement(
                    340,
                    @"[In Appendix B: Product Behavior][<14> Section 3.1.4.8:] The file [which contains the result of the import operation] includes a status code indicating the success of the operation or an error code. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(339, this.Site))
            {
                // If a status code is included in the file which is generated by ImportWeb, it means the file 
                // that contains the result of the export operation has been successfully generated. R339 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R339");

                // Verify MS-SITESS requirement: MS-SITESS_R339
                // Since the returned status code is not empty, this requirement can be captured directly.
                Site.CaptureRequirement(
                    339,
                    @"[In Appendix B: Product Behavior] <14> Section 3.1.4.8: Windows SharePoint Services creates a file that contains the result of the import operation in the server location specified in the request message. (Microsoft Office SharePoint Server 2007 and above follow this behavior.)");
            }
            #endregion Capture requirements

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ExportWeb operation when exporting a site whose 
        /// content exceeds 0x0018 megabytes, with cabSize parameter set in different values that 
        /// smaller than 0x18. In this case, multiple content migration package files 
        /// are expected to be exported.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC02_ExportingMutiplePackages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site), @"Test is executed only when R5311Enabled is set to true.");

            #region Variables
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.SpecialSubsiteUrl, this.Site);
            string exportUrlLowerCase = Common.GetConfigurationPropertyValue(Constants.SpecialSubsiteUrl, this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            int cabSize = 1;
            int exportWebResult = 0;
            string[] exportWebFilesWithCabSize1 = null;
            string[] expectedWebFilesWithCabSize1 = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            if (Common.IsRequirementEnabled(7411, this.Site))
            {
                // Invoke the ExportWeb operation with webUrl in lower case. Since The webUrl SHOULD be case-sensitive, 4 is expected to be returned.
                exportWebResult = this.sitessAdapter.ExportWeb(Constants.ExportJobName, exportUrlLowerCase, dataPath, true, true, true, cabSize);

                #region Capture requirements

                // If 4 is returned, MS-SITESS_R7411 can be captured.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R7411");

                // Verify MS-SITESS requirement: MS-SITESS_R7411
                Site.CaptureRequirementIfAreEqual<int>(
                    4,
                    exportWebResult,
                    7411,
                    @"[In ExportWeb] [webUrl:] Implementation does be case-sensitive.(Windows SharePoint Services 3.0 and above products follow this behavior.)");

                // If 4 is returned, MS-SITESS_R57 can be captured.
                this.VerifyExportWebErrorCode(exportWebResult);

                #endregion Capture requirements
            }

            // Invoke the ExportWeb operation with cabSize set to 1, 1 is expected to be returned and two cmp files are expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 3, this.Site, this.sutAdapter);
            exportWebFilesWithCabSize1 = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // As specified in MS-SITESS protocol section 3.1.4.2.2.1, if multiple content migration package files are created, 
            // then the server MUST also append a positive, incrementing integer to the name of each subsequent content migration package file.
            expectedWebFilesWithCabSize1 = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension,
                exportJobName + "1" + Constants.CmpExtension
            };

            #region Capture requirements

            // If exported files are consistent with the expected, it means multiple cmp files 
            // are created and a positive, incrementing integer is append to the second file's 
            // name (i.e. ExportWeb.cmp & ExportWeb1.cmp), R73 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R73: the exported file name list is {0}", files);

            // Verify MS-SITESS requirement: MS-SITESS_R73
            bool isVerifyR73 = AdapterHelper.CompareStringArrays(expectedWebFilesWithCabSize1, exportWebFilesWithCabSize1);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR73,
                73,
                @"[In ExportWeb] [jobName:] If multiple content migration package files are created, then the server MUST also append a positive, incrementing integer to the name of each subsequent content migration package file.");

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ExportWeb operation when exporting a site whose 
        /// content equals 0x0018 megabytes, with cabSize parameter set out of the valid range (0 and 0x400 respectively). 
        /// In this case, only one content migration package file is expected to be exported.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC03_ExportingEqualto0x18CabSize()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site), @"Test is executed only when R5311Enabled is set to true.");

            #region Variables
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.SpecialSubsiteUrl, this.Site);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            int cabSize = 0;
            int exportWebResult = 0;
            string[] exportWebFiles = null;
            string[] expectedWebFiles = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with cabSize set to 0, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            // If exported files are inconsistent with the expected, log it.
            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            this.sutAdapter.EmptyDocumentLibrary(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site));

            // Invoke the ExportWeb operation with cabSize set to 0x0400, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = 0;
            cabSize = 0x400;
            exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            // If exported files are inconsistent with the expected, log it.
            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            this.sutAdapter.EmptyDocumentLibrary(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site));

            // Invoke the ExportWeb operation with cabSize set to 0x0018, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = 0;
            cabSize = 0x18;
            exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            // If exported files are inconsistent with the expected, log it.
            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            #endregion Capture requirements

            this.sutAdapter.EmptyDocumentLibrary(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site));

            // Invoke the ExportWeb operation with cabSize set to 0x0018, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = 0;
            cabSize = -1;
            exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            if (Common.IsRequirementEnabled(532, this.Site))
            {
                bool isR532Verified = exportWebResult == 1 || exportWebResult == 7;

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R532");

                // Verify MS-SITESS requirement: MS-SITESS_R239
                Site.CaptureRequirementIfIsTrue(
                    isR532Verified,
                    532,
                    @"[In Appendix B: Product Behavior] <4> Section 3.1.4.2.2.1: If the value of cabSize is less than zero, Implementation does return a value of 1 or 7  but the server does not successfully complete the operation. The return code is not deterministic.(Windows SharePoint Services 3.0, SharePoint Foundation 2010, and SharePoint Foundation 2013 follow this behavior.)");
            }

        }

        /// <summary>
        /// This test case is designed to verify the ExportWeb operation when the ExportWebResult equals to InvalidExportUrl.
        /// </summary>s
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC04_ExportingFailureInvalidExportUrl()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site), @"Test is executed only when R5311Enabled is set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = siteUrl + "/" + Common.GetConfigurationPropertyValue(Constants.NonExistentSiteName, this.Site);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            int cabSize = 50;
            int exportWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with invalid webUrl, 4 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(Constants.ExportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 4 is returned, R90 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R90");

            // Verify MS-SITESS requirement: MS-SITESS_R90
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                exportWebResult,
                90,
                @"[In ExportWebResponse] If the value of ExportWebResult is 4, it specifies InvalidExportUrl: The site specified in the webUrl is not accessible.");

            // If 4 is returned, R57 can be captured.
            this.VerifyExportWebErrorCode(exportWebResult);

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ExportWeb operation when the ExportWebResult equals to ExportFileNoAccess.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC05_ExportingFailureExportFileNoAccess()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site), @"Test is executed only when R5311Enabled is set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string libraryName = Common.GetConfigurationPropertyValue(Constants.InvalidLibraryName, this.Site);
            string dataPath = siteUrl + "/" + libraryName;
            int cabSize = 50;
            int exportWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with invalid dataPath, 5 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(Constants.ExportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 5 is returned, R91 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R91");

            // Verify MS-SITESS requirement: MS-SITESS_R91
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                exportWebResult,
                91,
                @"[In ExportWebResponse] If the value of ExportWebResult is 5, it specifies ExportFileNoAccess: The location specified in the dataPath is not accessible.");

            // If 5 is returned, R57 can be captured.
            this.VerifyExportWebErrorCode(exportWebResult);

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ExportWeb operation when the ExportWebResult equals to ExportFileNoAccess.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC06_ExportingFailureOverwriteFailure()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site), @"Test is executed only when R5311Enabled is set to true.");

            #region Variables
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            int cabSize = 50;
            string[] exportWebFiles = null;
            int exportWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid dataPath, 1 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Invoke the ExportWeb operation with overwrite set to false and exportJobName set to an existing job name, 5 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, false, cabSize);

            #region Capture requirements
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R82");

            // Verify MS-SITESS requirement: MS-SITESS_R82
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                exportWebResult,
                82,
                @"[In ExportWeb] [overWrite:] The server MUST NOT overwrite existing file(s) with the new file(s) if false is specified.");

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R392");

            // Verify MS-SITESS requirement: MS-SITESS_R392
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                exportWebResult,
                392,
                @"[In ExportWeb] [overWrite:] If the server cannot create a new file due to this parameter set to FALSE, it MUST return error code 5 as defined in 3.1.4.2.2.2");

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ImportWeb operation when overwriting the log file to fail.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC07_ImportingFailureOverwriteFailure()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = siteUrl + "/" + this.newSubsite;
            string anotherImportUrl = siteUrl + "/" + Common.GetConfigurationPropertyValue(Constants.NonExistentSiteName, this.Site);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string logPath = string.Empty;
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string importJobName = Constants.ImportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { dataPath + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            string[] exportWebFiles = null;
            string[] importWebFiles = null;
            int exportWebResult = 0;
            int importWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid dataPath, 1 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            logPath = dataPath;

            // Invoke the ImportWeb operation with valid parameters, expect 1 is returned. The overwrite parameter is set to false to verify it doesn't mess anything up if no file to be overwrite.
            importWebResult = this.sitessAdapter.ImportWeb(importJobName, importUrl, dataFiles, logPath, true, false);

            #region Capture requirements

            // If 1 is returned, requirements related with ImportWeb pending can be verified.
            this.VerifyImportWebInProgress(importWebResult);

            #endregion Capture requirements

            files = TestSuiteHelper.VerifyExportAndImportFile(importJobName, 1, this.Site, this.sutAdapter);
            importWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            this.isWebImportedSuccessfully = true;

            // Invoke the ImportWeb operation with overwrite set to false, expect 11 is returned.
            importWebResult = this.sitessAdapter.ImportWeb(importJobName, anotherImportUrl, dataFiles, logPath, true, false);

            #region Capture requirements
            // If overwrite is set to false, R263 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R263");

            // Verify MS-SITESS requirement: MS-SITESS_R263
            Site.CaptureRequirementIfAreEqual<int>(
                11,
                importWebResult,
                263,
                @"[In ImportWeb] [overWrite:] The server MUST NOT overwrite existing files [at the location specified by logPath] with new files if false is specified.");

            // 11 will be returned if overwrite is set to false, so R393 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R393");

            // Verify MS-SITESS requirement: MS-SITESS_R393
            Site.CaptureRequirementIfAreEqual<int>(
                11,
                importWebResult,
                393,
                @"[In ImportWeb] [overWrite:] If the server cannot create a new file because this parameter is set to false, it MUST return error code 11 as specified in section 3.1.4.8.2.2.");
            #endregion Capture requirements

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ImportWeb operation when the ImportWebResult equals to InvalidImportUrl.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC08_ImportingFailureInvalidImportUrl()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables

            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = Common.GetConfigurationPropertyValue(Constants.NonExistentImportUrl, this.Site);
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string logPath = string.Empty;
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { dataPath + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            int exportWebResult = 0;
            int importWebResult = 0;
            string[] exportWebFiles = null;
            string[] expectedWebFiles = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid parameters, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Format the expected file names in the document library. Only one cmp file (i.e. ExportWeb.cmp) is expected to be exported in the document library.
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            logPath = dataPath;

            // Invoke the ImportWeb operation with invalid webUrl, 4 is expected to be returned.
            importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrl, dataFiles, logPath, true, true);

            #region Capture requirements

            // If 4 is returned, R269 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R269");

            // Verify MS-SITESS requirement: MS-SITESS_R269
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                importWebResult,
                269,
                @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 4, it specifies InvalidImportUrl: The site specified in the webUrl is not accessible.");

            this.VerifyImportWebErrorCode(importWebResult);

            #endregion Capture requirements

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ImportWeb operation when ImportWebResult equals to ImportFileNoAccess.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC09_ImportingFailureImportFileNoAccess()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5391Enabled is set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = siteUrl + "/" + this.newSubsite;
            string logPath = Common.GetConfigurationPropertyValue(Constants.SiteCollectionUrl, this.Site) + "/" + Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site);
            string[] dataFiles = new string[] { exportUrl + "/" + Common.GetConfigurationPropertyValue(Constants.InvalidLibraryName, this.Site) + "/" + Constants.ExportJobName + Constants.CmpExtension };
            int importWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ImportWeb operation with invalid dataFiles, 5 is expected to be returned.
            importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrl, dataFiles, logPath, true, true);

            #region Capture requirements

            this.VerifyImportWebErrorCode(importWebResult);

            // If 5 is returned, R270 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R270");

            // Verify MS-SITESS requirement: MS-SITESS_R270
            Site.CaptureRequirementIfAreEqual<int>(
                5,
                importWebResult,
                270,
                @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 5, it specifies ImportFileNoAccess: At least one location specified in dataFiles is not accessible.");

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();
            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ImportWeb operation when the ImportWebResult equals to LogFileNoAccess.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC10_ImportingFailureLogFileNoAccess()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = siteUrl + "/" + this.newSubsite;
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string logPath = string.Empty;
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { dataPath + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            int exportWebResult = 0;
            int importWebResult = 0;
            string[] exportWebFiles = null;
            string[] expectedWebFiles = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid parameters, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Format the expected file names in the document library. Only one cmp file (i.e. ExportWeb.cmp) is expected to be exported in the document library.
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            // Invoke the ImportWeb operation with invalid logPath, 11 is expected to returned.
            logPath = exportUrl + "/" + Common.GetConfigurationPropertyValue(Constants.InvalidLibraryName, this.Site);
            importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrl, dataFiles, logPath, true, true);

            #region Capture requirements

            this.VerifyImportWebErrorCode(importWebResult);

            // If 11 is returned, R273 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R273");

            // Verify MS-SITESS requirement: MS-SITESS_R273
            Site.CaptureRequirementIfAreEqual<int>(
                11,
                importWebResult,
                273,
                @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 11, it specifies LogFileNoAccess: The location specified by logPath is not accessible.");

            #endregion Capture requirements

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is designed to verify the ImportWeb operation when the ImportWebResult equals to ImportWebNotEmpty.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC11_ImportingFailureImportWebNotEmpty()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrlExported = string.Empty;
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string logPath = string.Empty;
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { exportUrl + "/" + Common.GetConfigurationPropertyValue(Constants.InvalidLibraryName, this.Site) + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            int exportWebResult = 0;
            int importWebResult = 0;
            string[] exportWebFiles = null;
            string[] expectedWebFiles = null;
            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid parameters, 1 is expected to be returned and only one cmp file is expected to be exported.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Format the expected file names in the document library. Only one cmp file (i.e. ExportWeb.cmp) is expected to be exported in the document library.
            expectedWebFiles = new string[]
            {
                exportJobName + Constants.SntExtension,
                exportJobName + Constants.CmpExtension
            };

            Site.Assert.IsTrue(AdapterHelper.CompareStringArrays(expectedWebFiles, exportWebFiles), "Export web files' names should be {0} and {1}, and actually the exported file list is: {2}", expectedWebFiles[0], expectedWebFiles[1], files);

            logPath = dataPath;
            importUrlExported = exportUrl;

            // Under SharePoint services 2007, the transport portion of the parameter webUrl should be lower case.
            if (importUrlExported.StartsWith("HTTP:", StringComparison.Ordinal))
            {
                importUrlExported = importUrlExported.Replace("HTTP:", "http:");
            }
            else if (importUrlExported.StartsWith("HTTPS:", StringComparison.Ordinal))
            {
                importUrlExported = importUrlExported.Replace("HTTPS:", "https:");
            }

            // As specified in MS-SITESS, the server name portion of the parameter webUrl should be lower case.
            string sutComputerName = Common.GetConfigurationPropertyValue(Constants.SutComputerName, this.Site);
            importUrlExported = importUrlExported.Replace(sutComputerName, sutComputerName.ToLower(System.Globalization.CultureInfo.CurrentCulture));

            // Invoke the ImportWeb operation with webUrl set to the site just exported, which is not a blank site, 8 is expected to be returned.
            importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrlExported, dataFiles, logPath, true, true);

            #region Capture requirements

            this.VerifyImportWebErrorCode(importWebResult);

            // If 8 is returned, R272 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R272");

            // Verify MS-SITESS requirement: MS-SITESS_R272
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                importWebResult,
                272,
                @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 8, it specifies ImportWebNotEmpty: The location specified by webUrl corresponds to an existing site that is not a blank site.");

            #endregion Capture requirements

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// This test case is used to verify the ImportWeb operation when the logPath is omitted or empty.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S01_TC12_ImportingFailureLogPathEmpty()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5311, this.Site) && Common.IsRequirementEnabled(5391, this.Site), @"Test is executed only when R5311Enabled and R5391Enabled are set to true.");

            #region Variables
            string siteUrl = Common.GetConfigurationPropertyValue(Constants.SiteUrl, this.Site);
            string exportUrl = Common.GetConfigurationPropertyValue(Constants.NormalSubsiteUrl, this.Site);
            string importUrl = siteUrl + "/" + this.newSubsite;
            string dataPath = Common.GetConfigurationPropertyValue(Constants.DataPath, this.Site);
            string exportJobName = Constants.ExportJobName + Common.FormatCurrentDateTime();
            string[] dataFiles = new string[] { dataPath + "/" + exportJobName + Constants.CmpExtension };
            int cabSize = 50;
            string[] exportWebFiles = null;
            string[] exportWebAndImportWebFiles = null;
            int exportWebResult = 0;
            int importWebResult = 0;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportWeb operation with valid dataPath, 1 is expected to be returned.
            exportWebResult = this.sitessAdapter.ExportWeb(exportJobName, exportUrl, dataPath, true, true, true, cabSize);

            #region Capture requirements

            // If 1 is returned, requirements related with ExportWeb pending can be verified.
            this.VerifyExportWebInProgress(exportWebResult);

            #endregion Capture requirements

            string files = TestSuiteHelper.VerifyExportAndImportFile(exportJobName, 2, this.Site, this.sutAdapter);
            exportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            // Invoke the ImportWeb operation with valid parameters, expect 1 is returned. The logPath is set to null to verify the server MUST NOT create any files describing the progress or status of the ImportWeb operation.
            importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrl, dataFiles, null, true, true);

            #region Capture requirements

            // If 1 is returned, requirements related with ImportWeb pending can be verified.
            this.VerifyImportWebInProgress(importWebResult);

            #endregion Capture requirements

            
            files = TestSuiteHelper.VerifyExportAndImportFile(null, 2, this.Site, this.sutAdapter);
            exportWebAndImportWebFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

            bool isExportImportWebFiles = exportWebAndImportWebFiles.ToString().Contains(Constants.ImportJobName + Constants.SntExtension);

            #region Capture requirements

            // If isExportImportWebFiles is false, R409 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R409");

            // Verify MS-SITESS requirement: MS-SITESS_R409
            Site.CaptureRequirementIfIsFalse(
                isExportImportWebFiles,
                409,
                @"[In ImportWeb] [logPath:] If this element is omitted, the server MUST NOT create any files describing the progress or status of the operation.");

            #endregion Capture requirements
            

            bool isErrorOccured = false;
            SoapException soapException = null;
            if (Common.IsRequirementEnabled(391,this.Site)|| Common.IsRequirementEnabled(271001, this.Site))
            {
                try
                {
                    // Invoke the ImportWeb operation with valid parameters, expect an error is returned. The logPath is set to empty string.
                    importWebResult = this.sitessAdapter.ImportWeb(Constants.ImportJobName, importUrl, dataFiles, string.Empty, true, true);
                }
                catch (SoapException ex)
                {
                    soapException = ex;
                    isErrorOccured = true;
                }
                
                if (Common.IsRequirementEnabled(391, this.Site))
                {
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R391");

                    // Verify MS-SITESS requirement: MS-SITESS_R391
                    Site.CaptureRequirementIfIsTrue(
                        isErrorOccured,
                        391,
                        @"[In ImportWeb] [logPath:] If this element is omitted, the server MUST NOT create any files describing the progress or status of the operation.");

                }

                if (Common.IsRequirementEnabled(271001, this.Site))
                {
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R271001");

                    // Verify MS-SITESS requirement: MS-SITESS_R271001
                    Site.CaptureRequirementIfIsTrue(
                        importWebResult==7,
                        271001,
                        @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 7, it specifies ImportWebNotEmpty: The site specified in the webUrl cannot be created.");
                }               
            }

            this.isWebImportedSuccessfully = true;

            #region Capture requirements

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            this.VerifyOperationExportWeb();

            // Verify that Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            this.VerifyOperationImportWeb();

            #endregion Capture requirements
        }

        /// <summary>
        /// If the value of ExportWebResult is 1, it specifies Pending: The operation is in progress. R58 and R89 can be captured.
        /// </summary>
        /// <param name="actualValue">Value of ExportWebResult</param>
        public void VerifyExportWebInProgress(int actualValue)
        {
            // If 1 is returned, R58 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R58");

            // Verify MS-SITESS requirement: MS-SITESS_R58
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                actualValue,
                58,
                @"[In ExportWeb] Upon the start of the export operation, a result of 1 MUST be sent as the response indicating that the operation is in progress.");

            // If the files are equal to the expected file, the server process this operation and the result is ok, then R89 can be verified.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R89");

            // Verify MS-SITESS requirement: MS-SITESS_R89
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                actualValue,
                89,
                @"[In ExportWebResponse] If the value of ExportWebResult is 1, it specifies Pending: The operation is in progress.");
        }

        /// <summary>
        /// If ExportWeb failed, errorCode should be 4, 5, 6, 7, or 8. R57 can be captured.
        /// </summary>
        /// <param name="errorCodeValue">The error code returned by the server.</param>
        public void VerifyExportWebErrorCode(int errorCodeValue)
        {
            // If expectedValue is returned as 4, 5, 6, 7 or 8, R57 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R57, the value of error code is {0}", errorCodeValue);

            // Verify MS-SITESS requirement: MS-SITESS_R57
            bool isVerifyR57 = 4 == errorCodeValue || 5 == errorCodeValue || 6 == errorCodeValue || 7 == errorCodeValue || 8 == errorCodeValue;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR57,
                57,
                @"[In ExportWeb] In cases of permission, space restrictions or other conditions that prevent the execution of the operation, an error code MUST be included in the response message as specified in section 3.1.4.2.2.2 [error code list: 4, 5, 6, 7, 8].");
        }

        /// <summary>
        /// If the value of ImportWebResult is 1, it specifies Pending: The operation is in progress. R239 and R267 can be captured.
        /// </summary>
        /// <param name="actualValue">value of ImportWebResult</param>
        public void VerifyImportWebInProgress(int actualValue)
        {
            // If 1 is returned, R239 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R239");

            // Verify MS-SITESS requirement: MS-SITESS_R239
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                actualValue,
                239,
                @"[In ImportWeb] Upon the start of the import operation, a result of 1 MUST be sent as the response indicating that the operation is in progress.");

            // If no error means the server process this operation and the result is ok, then we verify this requirement.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R267");

            // Verify MS-SITESS requirement: MS-SITESS_R267
            Site.CaptureRequirementIfAreEqual<int>(
                1,
                actualValue,
                267,
                @"[In ImportWebResponse] [ImportWebResult:] If the value of ImportWebResult is 1, it specifies Pending: The operation is in progress.");
        }

        /// <summary>
        /// If ImportWeb failed, errorCode should be 2, 4, 5, 6, 8 or 11. R238 can be captured.
        /// </summary>
        /// <param name="errorCodeValue">Value of errorCode</param>
        public void VerifyImportWebErrorCode(int errorCodeValue)
        {
            // If errorCode is 2, 4, 5, 6, 8 or 11, R238 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R238, the value of error code is {0}", errorCodeValue);

            // Verify MS-SITESS requirement: MS-SITESS_R238
            bool isVerifyR238 = 2 == errorCodeValue || 4 == errorCodeValue || 5 == errorCodeValue || 6 == errorCodeValue || 8 == errorCodeValue || 11 == errorCodeValue;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR238,
                238,
                @"[In ImportWeb] If a condition occurs that prevents the server from executing the operation, an error code MUST be included in the response message as specified in section 3.1.4.8.2.2 [error code list: 2, 4, 5, 6, 8, 11].");
        }

        /// <summary>
        /// This method is used to verify Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
        /// </summary>
        public void VerifyOperationExportWeb()
        {
            // If code can run to here, it means Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5311, Microsoft Windows SharePoint Services 3.0 and above support operation ExportWeb.");

            // Verify MS-SITESS requirement: MS-SITESS_R5311
            Site.CaptureRequirement(
                5311,
                @"[In Appendix B: Product Behavior] <2> Section 3.1.4.2: Implementation does support this operation [ExportWeb].(Windows SharePoint Services 3.0 and above follow this behavior.)");
        }

        /// <summary>
        /// This method is used to verify Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
        /// </summary>
        public void VerifyOperationImportWeb()
        {
            // If code can run to here, it means Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5391, Microsoft Windows SharePoint Services 3.0 and above support operation ImportWeb.");

            // Verify MS-SITESS requirement: MS-SITESS_R5391
            Site.CaptureRequirement(
                5391,
                @"[In Appendix B: Product Behavior] <13> Section 3.1.4.8: Implementation does support this operation [ImportWeb].(Windows SharePoint Services 3.0 and above follow this behavior.)");
        }

        #endregion Scenario 1 Migrate a site

        #endregion Test Cases

        #region Test Case Initialization & Cleanup

        /// <summary>
        /// Test Case Initialization.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.sitessAdapter = Site.GetAdapter<IMS_SITESSAdapter>();
            Common.CheckCommonProperties(this.Site, true);
            this.sutAdapter = Site.GetAdapter<IMS_SITESSSUTControlAdapter>();
            this.newSubsite = "NewSubsite" + Common.FormatCurrentDateTime();
        }

        /// <summary>
        /// Test Case Cleanup.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            // Remove all files in a document Library, which is used as the store location for the files exported.
            this.sutAdapter.EmptyDocumentLibrary(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, this.Site));

            // Remove an imported site if it is imported successfully.
            if (this.isWebImportedSuccessfully)
            {
                // Since the DeleteWeb operation is not supported under the following product, an SUT method is called to do the same thing.
                if (Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.SharePointServer2007, System.StringComparison.CurrentCultureIgnoreCase)
                    || Common.GetConfigurationPropertyValue(Constants.SutVersion, this.Site).Equals(Constants.WindowsSharePointServices3, System.StringComparison.CurrentCultureIgnoreCase))
                {
                    this.sutAdapter.RemoveWeb(this.newSubsite);
                }
                else
                {
                    try
                    {
                        this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);
                        this.sitessAdapter.DeleteWeb(Common.GetConfigurationPropertyValue(Constants.SiteName, this.Site) + "/" + this.newSubsite);
                    }
                    catch (SoapException e)
                    {
                        Site.Log.Add(LogEntryKind.Comment, "S1_MigrateSite_TestCleanup: ");
                        Site.Log.Add(LogEntryKind.Comment, e.Code.ToString());
                        Site.Log.Add(LogEntryKind.Comment, e.Message);
                    }
                }

                this.isWebImportedSuccessfully = false;
            }

            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion
    }
}