//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Runtime.Remoting;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class contains test case definitions related to ExportSolution operation.
    /// </summary>
    [TestClass]
    public class S04_ExportSolution : TestClassBase
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
        /// The name of a solution package which is to be exported, which is a compressed file that can be deployed to a server farm or a site.
        /// </summary>
        private string solutionName = string.Empty;

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

        #endregion Test Suite Initialization & Cleanup

        #region Test Cases

        #region Scenario 4 ExportSolution

        /// <summary>
        /// This test case is designed to verify the successful status of ExportSolution.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S04_TC01_ExportSolutionSucceed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(5301, this.Site), @"Test is executed only when R5301Enabled is set to true.");

            #region Variables
            string exportResult = string.Empty;
            string galleryName = Common.GetConfigurationPropertyValue(Constants.SolutionGalleryName, this.Site);
            string[] exportFiles = null;
            string[] expectedFiles = null;

            #endregion Variables

            // Initialize the web service with an authenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Authenticated);

            // Invoke the ExportSolution operation, a url is expected to be returned and a solution file is expected to be exported.
            exportResult = this.sitessAdapter.ExportSolution(this.solutionName + Constants.WspExtension, Constants.SolutionTitle, Constants.SolutionDescription, true, true);

            DateTime beforeLoop = DateTime.Now;
            TimeSpan waitTime = new TimeSpan();
            int repeatTime = 0;
            int totalRepeatTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ExportRepeatTime, this.Site));
            string solutions = string.Empty;
            do
            {
                // It is assumed that the server will generate all the exported files and log file after the preconfigured time period.
                int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ExportWaitTime, this.Site));
                Thread.Sleep(1000 * sleepTime);

                // Get all the file names in the solution library.
                solutions = this.sutAdapter.GetDocumentLibraryFileNames(string.Empty, string.Empty, galleryName, this.solutionName);
                exportFiles = solutions == null ? null : solutions.TrimEnd(new char[] { ';' }).Split(';');

                // If the expected files are created, the operation is considered as succeed.
                if (exportFiles != null && exportFiles.Length == 1)
                {
                    break;
                }

                // If the server could not generate all the exported files and log file after the time period ExportWaitTime for some unknown reasons, for example, limit of server resources, try to repeat this sequence again.
                repeatTime++;
            }
            while (repeatTime < totalRepeatTime);
            waitTime = DateTime.Now - beforeLoop;

            // If the server still does not generate all the exported files and log file after repeating, or the server generate unexpected number of files, the operation is considered as failed.
            if (exportFiles == null)
            {
                Site.Assert.Fail("The server does not export the files after {0} seconds", waitTime);
            }
            else if (exportFiles.Length != 1)
            {
                Site.Assert.Fail("The server does not export {0} files as expected but exports {1} file{2} in actual, and the exported file name list is {3}", 1, exportFiles.Length, exportFiles.Length > 1 ? "s" : string.Empty, solutions);
            }

            // Format the expected file names in the solution gallery, only one solution file (i.e. SolutionName) is expected.
            expectedFiles = new string[] { this.solutionName + Constants.WspExtension };

            // If returned value is not a url or exported files are inconsistent with the expected, log it.
            Site.Assert.IsTrue(
                Uri.IsWellFormedUriString(exportResult, UriKind.Relative),
                "ExportSolution should return a valid Uri, actual uri {0}.",
                exportResult);

            Site.Assert.IsTrue(
                AdapterHelper.CompareStringArrays(expectedFiles, exportFiles),
                "ExportSolution should export the solution file.");

            // If returned value is a url and exported files are consistent with the expected, it means the ExportSolution operation succeed.
            // Invoke the ExportSolution operation again, a url is expected to be returned and a second solution file is expected to be exported.
            exportResult = this.sitessAdapter.ExportSolution(this.solutionName + Constants.WspExtension, Constants.SolutionTitle, Constants.SolutionDescription, true, true);

            beforeLoop = DateTime.Now;
            waitTime = new TimeSpan();
            repeatTime = 0;
            solutions = string.Empty;
            do
            {
                // It is assumed that the server will generate all the exported files and log file after the preconfigured time period.
                int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ExportWaitTime, this.Site));
                Thread.Sleep(1000 * sleepTime);

                // Get all the file names in the solution library.
                solutions = this.sutAdapter.GetDocumentLibraryFileNames(string.Empty, string.Empty, galleryName, this.solutionName);
                exportFiles = solutions == null ? null : solutions.TrimEnd(new char[] { ';' }).Split(';');

                // If the expected files are created, the operation is considered as succeed.
                if (exportFiles != null && exportFiles.Length == 2)
                {
                    break;
                }

                // If the server could not generate all the exported files and log file after the time period ExportWaitTime for some unknown reasons, for example, limit of server resources, try to repeat this sequence again.
                repeatTime++;
            }
            while (repeatTime < totalRepeatTime);
            waitTime = DateTime.Now - beforeLoop;

            // If the server still does not generate all the exported files and log file after repeating, or the server generate unexpected number of files, the operation is considered as failed.
            if (exportFiles == null)
            {
                Site.Assert.Fail("The server does not export the files after {0} seconds", waitTime);
            }
            else if (exportFiles.Length != 2)
            {
                Site.Assert.Fail("The server does not export {0} files as expected but exports {1} file{2} in actual, and the exported file name list is {3}", 2, exportFiles.Length, exportFiles.Length > 1 ? "s" : string.Empty, solutions);
            }

            // Format the expected file names in the solution gallery, two solution files (i.e. SolutionName & SolutionName2) are expected.
            expectedFiles = new string[]
                {
                        this.solutionName + Constants.WspExtension,
                        this.solutionName + "2" + Constants.WspExtension
                };

            // If returned value is not a url or exported files are inconsistent with the expected, log it.
            Site.Assert.IsTrue(
                Uri.IsWellFormedUriString(exportResult, UriKind.Relative),
                "ExportSolution should return a valid Uri, actual uri {0}.",
                exportResult);

            #region Capture requirements

            // If exported files are consistent with the expected, it means multiple wsp files 
            // are created and a positive, incrementing integer is append to the second file's 
            // name (i.e. ExportSolution.wsp & ExportSolution2.wsp), so R384 can be captured.
            bool isMutipleFile = AdapterHelper.CompareStringArrays(expectedFiles, exportFiles);

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R384, the export file is {0}", exportFiles[1]);

            // Verify MS-SITESS requirement: MS-SITESS_R384
            Site.CaptureRequirementIfIsTrue(
                isMutipleFile,
                384,
                @"[In ExportSolution] [solutionFileName:] If a solution with the specified name already exists in the solution gallery, the server retry with <filename>2.wsp, where <filename> is obtained from solutionFileName after excluding the extension.");

            // If exported files are consistent with the expected, it means multiple wsp files 
            // are created and a positive, incrementing integer is append to the second file's 
            // name (i.e. ExportSolution.wsp & ExportSolution2.wsp), so R38 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R38, the export file is {0}", exportFiles[1]);

            // Verify MS-SITESS requirement: MS-SITESS_R38
            Site.CaptureRequirementIfIsTrue(
                isMutipleFile,
                38,
                @"[In ExportSolution] [solutionFileName:] If a unique name is obtained, the server MUST continue with that name [to create a solution file using this unique name].");

            // If code can run to here, it means that Microsoft SharePoint Foundation 2010 and above support method ExportSolution.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R5301, Microsoft SharePoint Foundation 2010 and above support method ExportSolution.");

            // Verify MS-SITESS requirement: MS-SITESS_R5301
            Site.CaptureRequirement(
                5301,
                @"[In Appendix B: Product Behavior] Implementation does support this method [ExportSolution]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            #endregion Capture requirements
        }

        #endregion Scenario 4 ExportSolution

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
            this.solutionName = "NewSolution" + Common.FormatCurrentDateTime();
        }

        /// <summary>
        /// Test Case Cleanup.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            if (Common.IsRequirementEnabled(5301, this.Site))
            {
                DateTime beforeLoop = DateTime.Now;
                TimeSpan waitTime = new TimeSpan();
                double timeOut = Convert.ToDouble(Common.GetConfigurationPropertyValue(Constants.ExportWaitTime, this.Site));

                // Delete the exported solutions. Since when the ExportSolution operation is not completed, try to remove the exported solutions will cause RemotingException, 
                // and no other methods to check if the operation is finished, use "do while" loop for deleting the solutions.
                bool deleteResult = false;
                do
                {
                    try
                    {
                        deleteResult = this.sutAdapter.RemoveAllSolution(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.SolutionGalleryName, this.Site));
                    }
                    catch (RemotingException e)
                    {
                        Site.Log.Add(LogEntryKind.Comment, "Delete the exported solutions failed: " + e.Message);
                    }

                    waitTime = DateTime.Now - beforeLoop;
                }
                while ((!deleteResult) && waitTime.TotalSeconds < timeOut);

                if (!deleteResult)
                {
                    Site.Assert.Fail("The server does not delete the exported solutions after {0} seconds", timeOut);
                }
            }

            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion Test Case Initialization & Cleanup
    }
}