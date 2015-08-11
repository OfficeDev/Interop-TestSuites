namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that contains the helper methods used by MS-VERSS test cases.
    /// </summary>
    public class TestSuiteHelper
    {
        /// <summary>
        /// Verify exported files and log file for the ExportWeb, ImportWeb, ExportSolution and ExportWorkflowTemplate operations.
        /// </summary>
        /// <param name="fileName">A string indicates the expected file name.</param>
        /// <param name="fileNumber">A Integer indicates the expected file number.</param>
        /// <param name="testSite">The instance of ITestSite.</param>
        /// <param name="sutAdapter">The instance of the SUT control adapter.</param>
        /// <returns>A string indicates the exported filename.</returns>
        public static string VerifyExportAndImportFile(string fileName, int fileNumber, ITestSite testSite, IMS_SITESSSUTControlAdapter sutAdapter)
        {
            string[] exportFiles = null;
            DateTime beforeLoop = DateTime.Now;
            TimeSpan waitTime = new TimeSpan();
            int repeatTime = 0;
            int totalRepeatTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ExportRepeatTime, testSite));
            string files = string.Empty;
            do
            {
                // It is assumed that the server will generate all the exported files and log file after the preconfigured time period.
                int sleepTime = Convert.ToInt32(Common.GetConfigurationPropertyValue(Constants.ExportWaitTime, testSite));
                Thread.Sleep(1000 * sleepTime);

                // Get all the file names in the document library.
                files = sutAdapter.GetDocumentLibraryFileNames(string.Empty, string.Empty, Common.GetConfigurationPropertyValue(Constants.ValidLibraryName, testSite), fileName);
                exportFiles = files == null ? null : files.TrimEnd(new char[] { ';' }).Split(';');

                // If the expected files are created, the operation is considered as succeed.
                if (exportFiles != null && exportFiles.Length == fileNumber)
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
                testSite.Assert.Fail("The server does not export the files after {0} seconds", waitTime);
            }
            else if (exportFiles.Length != fileNumber)
            {
                testSite.Assert.Fail("The server does not export {0} files as expected but exports {1} file{2} in actual, and the exported file name list is {3}", 2, exportFiles.Length, exportFiles.Length > 1 ? "s" : string.Empty, files);
            }

            return files;
        }
    }
}