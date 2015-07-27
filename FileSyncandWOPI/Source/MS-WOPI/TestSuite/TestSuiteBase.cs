//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test Suite bas class contain the basic initialization/clean up logic and common helper methods.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets or sets the IMS_WOPIAdapter instance.
        /// </summary>
        protected static IMS_WOPIAdapter WopiAdapter { get; set; }

        /// <summary>
        /// Gets or sets the IMS_WOPISUTControlAdapter instance.
        /// </summary>
        protected static IMS_WOPISUTControlAdapter SutController { get; set; }

        /// <summary>
        /// Gets or sets a IMS_WOPISUTManageCodeControlAdapter type instance.
        /// </summary>
        protected static IMS_WOPIManagedCodeSUTControlAdapter WopiSutManageCodeControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets the document library name.
        /// </summary>
        protected static string TargetDocLibraryListName { get; set; }
        
        /// <summary>
        /// Gets or sets the test client name which the test suite runs on.
        /// </summary>
        protected static string CurrentTestClientName { get; set; }

        /// <summary>
        /// Gets or sets the file name counter which is used for per test cases.
        /// </summary>
        protected static uint FileNameCounterOfPerTestCases { get; set; }
 
        /// <summary>
        /// Gets or sets the absolute URL of a uploaded file which is located on WOPI SUT.
        /// </summary>
        protected static string UploadedFileUrl { get; set; }

        /// <summary>
        /// Gets or sets a collection which records the added files' absolute URLs. It is used to clean up all added files.
        /// </summary>
        private static List<string> AddedFilesRecorder { get; set; }
        #endregion 

        #region Test Suite Initialization

        /// <summary>
        /// Class level initialization method. 
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);

            // Execute the MS-WOPI initialization
            if (!TestSuiteHelper.VerifyRunInSupportProducts(TestSuiteBase.BaseTestSite))
            {
                return;
            }

            TestSuiteHelper.InitializeTestSuite(TestSuiteBase.BaseTestSite);
            WopiAdapter = TestSuiteHelper.WOPIProtocolAdapter;
            SutController = TestSuiteHelper.WOPISutControladapter;
            WopiSutManageCodeControlAdapter = TestSuiteHelper.WOPIManagedCodeSUTControlAdapter;
            CurrentTestClientName = TestSuiteHelper.CurrentTestClientName;

            if (string.IsNullOrEmpty(UploadedFileUrl))
            {
                UploadedFileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", TestSuiteBase.BaseTestSite);
            }

            if (string.IsNullOrEmpty(TargetDocLibraryListName))
            {
                TargetDocLibraryListName = Common.GetConfigurationPropertyValue("MSWOPIDocLibraryName", TestSuiteBase.BaseTestSite);
            }

            if (null == AddedFilesRecorder)
            {
                AddedFilesRecorder = new List<string>();
            }
        }

        /// <summary>
        /// Use ClassCleanup to run code after all tests in a class have run
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteHelper.CleanUpDiscoveryProcess(CurrentTestClientName, SutController);
            string discoveryListenerLog = DiscoveryRequestListener.GetLogs(typeof(DiscoveryRequestListener));
            BaseTestSite.Log.Add(LogEntryKind.Debug, discoveryListenerLog);

            // Clean up all the added files.
            DeleteCollectedFiles(AddedFilesRecorder);
            TestClassBase.Cleanup();
        }

        #endregion 

        #region Test Case Initialization

        /// <summary>
        /// Test case level initialization method. This method will be invoked for each test cases.
        /// </summary>
        [TestInitialize]
        public void MSWOPITestCaseInitialize()
        {
           this.InitializeCounterForPerTestCase();
           TestSuiteHelper.PerformSupportProductsCheck(this.Site);
        }

        #endregion 

        #region Helper method

        /// <summary>
        /// A method is used to get a unique file name for per test cases.
        /// </summary>
        /// <param name="isMultipleResourcesPerCase">A parameter represents bool value indicating this method is used for scenario "multiple resources per case". The default value is false.</param>
        /// <returns>A return value represents the unique file name.</returns>
        protected string GetUniqueFileName(bool isMultipleResourcesPerCase = false)
        {
            return this.GetUniqueResourceName("File", isMultipleResourcesPerCase);
        }

        /// <summary>
        /// A method is used to get a unique file name for "PutRelativeFile" usage.
        /// </summary>
        /// <returns>A return value represents the unique file name for "PutRelativeFile" usage.</returns>
        protected string GetUniqueFileNameForPutRelatived()
        {
            return this.GetUniqueResourceName("NewAddedByPutRelativeFile", true);
        }

        /// <summary>
        /// A method is used to add a file with unique file name into the specified document library list, and record the file URL into the AddedFilesRecorder.
        /// </summary>
        /// <param name="isMultipleResourcesPerCase">A parameter represents bool value indicating this method is used for scenario "multiple resources per case". The default value is false.</param>
        /// <returns>A return value represents the absolute URL of the added file.</returns>
        protected string AddFileToSUT(bool isMultipleResourcesPerCase = false)
        {
            string fileName = this.GetUniqueFileName(isMultipleResourcesPerCase);
            string addedFileUrl = SutController.AddFileToSUT(TargetDocLibraryListName, fileName);
            if (string.IsNullOrEmpty(addedFileUrl))
            {
                this.Site.Assert.Fail(
                            "Could not upload the file[{0}] to the Document library[{1}].",
                            fileName,
                            TargetDocLibraryListName);
            }

            string errorOfValidateFileUrl;
            if (!TryVerifyFileUrl(addedFileUrl, out errorOfValidateFileUrl))
            {
                this.Site.Assert.Fail(
                             "The Sut controller adapter does not return a valid file URL for uploaded the file[{0}] to the Document library[{1}].\r\nReturned URL:[{2}]\r\n validate error:[{3}]",
                             fileName,
                             TargetDocLibraryListName,
                             addedFileUrl,
                             errorOfValidateFileUrl);
            }

            AddedFilesRecorder.Add(addedFileUrl);
            return addedFileUrl;
        }

        /// <summary>
        /// A method is used to initialize the resource's counters for per test case.
        /// </summary>
        protected void InitializeCounterForPerTestCase()
        {
            FileNameCounterOfPerTestCases = 0;
        }
 
        /// <summary>
        /// A method is used to get the headers value from a web exception's http response.
        /// </summary>
        /// <param name="webException">A parameter represents the web exception instance.</param>
        /// <returns>A return value represents the HTTP response contain the error information which is contained in the web exception.</returns>
        protected virtual HttpWebResponse GetErrorResponseFromWebException(WebException webException)
        {
            if (null == webException)
            {
                throw new ArgumentNullException("webException");
            }

            HttpWebResponse errorResponse = webException.Response as HttpWebResponse;
            if (null == errorResponse)
            {
                string errorMsg = string.Format(
                                              "Could not parse the error response from the WebException when {0} exception happen.",
                                               webException.Message);
                throw new InvalidCastException(errorMsg);
            }
            else
            {
                return errorResponse;
            }
        }

        /// <summary>
        /// A method is used to collect added file absolute path by specified file name. After collecting the file path, the test suite will remove the collected file in test suite clean up process.
        /// </summary>
        /// <param name="normaUrlOfRelatedFile">A parameter represents the related file URL, which is used to locate the new added file's absolute path. The related file must be in same level location of the new added file.</param>
        /// <param name="newAddedFileName">A return value represents the file name of the new added file.</param>
        protected void CollectNewAddedFileForPutRelativeFile(string normaUrlOfRelatedFile, string newAddedFileName)
        {
            #region check parameter
            if (string.IsNullOrEmpty(newAddedFileName))
            {
                throw new ArgumentNullException("newAddedFileName");
            }

            if (string.IsNullOrEmpty(normaUrlOfRelatedFile))
            {
                throw new ArgumentNullException("normaUrlOfRelatedFile");
            }

            VerifyFileUrl(normaUrlOfRelatedFile);
            #endregion 

            Uri relatedFilePath = new Uri(normaUrlOfRelatedFile);
            string directoryNameValue = Path.GetDirectoryName(relatedFilePath.LocalPath);
            string newAddedFileLocalPath = Path.Combine(directoryNameValue, newAddedFileName);
            newAddedFileLocalPath = newAddedFileLocalPath.Replace(@"\", @"/");
            Uri newFilePath;
            if (!Uri.TryCreate(newAddedFileLocalPath, UriKind.Relative, out newFilePath))
            {
                throw new InvalidOperationException(string.Format("Could not combine the valid path for the new added file. Current:[{0}]", newAddedFileLocalPath));
            }

            string absolutePathValueOfNewFile = string.Format(
                                                        @"{0}://{1}{2}",
                                                        relatedFilePath.Scheme,
                                                        relatedFilePath.Host,
                                                        newFilePath.OriginalString);

            // Verify the new construct file URL.
            VerifyFileUrl(absolutePathValueOfNewFile);
            AddedFilesRecorder.Add(absolutePathValueOfNewFile);
        }

        /// <summary>
        /// A method is used to exclude a new added file's URL from the clean up process. Any files added by "AddFileToSUT" method will be deleted automatically in clean up process. This method will exclude a file from the clean up process.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the file URL, which is excluded from the clean up process.</param>
        protected void ExcludeFileFromTheCleanUpProcess(string fileUrl)
        {
            if (string.IsNullOrEmpty(fileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            VerifyFileUrl(fileUrl);
            if (null == AddedFilesRecorder || 0 == AddedFilesRecorder.Count)
            {
                throw new InvalidOperationException("There is no any added file in added files recorder.");
            }

            int indexOfExistFileUrl = AddedFilesRecorder.FindIndex(expectedUrlItme => expectedUrlItme.Equals(fileUrl, StringComparison.OrdinalIgnoreCase));
            if (indexOfExistFileUrl >= 0)
            {
                AddedFilesRecorder.RemoveAt(indexOfExistFileUrl);
            }
            else
            {
                throw new InvalidOperationException(string.Format("The expected file with URL[{0}] does not exist in added files recorder.", fileUrl));
            }
        }

        /// <summary>
        /// This method is used to get the status code value from web exception in an error Http response.
        /// </summary>
        /// <param name="response">A parameter represents the error HTTP response.</param>
        /// <returns>A return value represents the status code value.</returns>
        protected int GetStatusCodeFromHTTPResponse(HttpWebResponse response)
        {
            if (null == response)
            {
                throw new ArgumentNullException("response");
            }

            int statusCode = (int)response.StatusCode;
            this.Site.Log.Add(
                         LogEntryKind.Debug,
                         "The protocol server return a status code[{0}].",
                         statusCode);

            return statusCode;
        }

        /// <summary>
        /// A method used to delete collected files. If a collected file URL is not a valid file URL, this method will ignore it. If not all the valid file URLs are processed successfully, this method will raise an InvalidOperationException.
        /// </summary>
        /// <param name="collectedFiles">A parameter represents the collected file URLs.</param>
        private static void DeleteCollectedFiles(List<string> collectedFiles)
        {
            if (null == collectedFiles || 0 == collectedFiles.Count)
            {
                BaseTestSite.Log.Add(LogEntryKind.Debug, "There are no added files, skip the delete process.");
                return;
            }

            StringBuilder invalidUrlRecorder = new StringBuilder();
            StringBuilder validUrlRecorder = new StringBuilder();
            StringBuilder logsForValidUrlRecorder = new StringBuilder();
            foreach (string fileUrlItem in collectedFiles)
            {
                if (string.IsNullOrEmpty(fileUrlItem))
                {
                    continue;
                }

                string errorMsg;
                if (!TryVerifyFileUrl(fileUrlItem, out errorMsg))
                {
                    invalidUrlRecorder.AppendLine(string.Format(@"Invalid file URL:[{0}], Error:[{1}]", fileUrlItem, errorMsg));
                    continue;
                }
                
                // Construct the SUT controller input parameter.
                validUrlRecorder.Append(fileUrlItem + ",");

                // Log the files the test suite plan to delete.
                string fileName = TestSuiteHelper.GetFileNameFromFullUrl(fileUrlItem);
                logsForValidUrlRecorder.AppendLine(string.Format(@"File name:[{0}] File URL:[{1}]", fileName, fileUrlItem));
            }

            if (invalidUrlRecorder.Length != 0)
            {
                BaseTestSite.Log.Add(
                    LogEntryKind.Debug,
                    "There are some invalid URLs for collected files, test suite will skip these file URLs:\r\n{0}",
                    invalidUrlRecorder.ToString());
            }

            BaseTestSite.Log.Add(
                LogEntryKind.Debug,
                "Test suite prepare to delete these collected files:\r\n{0}",
                logsForValidUrlRecorder.ToString());

            // Call SUT controller method to delete files.
            string uploadedfilesUrls = validUrlRecorder.ToString(0, validUrlRecorder.Length - 1);
            bool areFilesDeletedSuccessful = false;
            try
            {
                areFilesDeletedSuccessful = SutController.DeleteUploadedFilesOnSUT(TargetDocLibraryListName, uploadedfilesUrls);
            }
            finally
            {
                AddedFilesRecorder.Clear();
            }

            if (!areFilesDeletedSuccessful)
            {
                throw new InvalidOperationException("Not all the collected files are deleted successfully.");
            }
        }
 
        /// <summary>
        /// A method used to verify a file URL whether is a valid file URL.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the file URL which will be validated.</param>
        /// <param name="errorForInvalidUrl">A parameter represents the validate error for the file URL. If the URL is a valid file URL, the out value of this parameter is empty.</param>
        /// <returns>Return 'true' indicating the specified URL is a valid file URL.</returns>
        private static bool TryVerifyFileUrl(string fileUrl, out string errorForInvalidUrl)
        {
            if (string.IsNullOrEmpty(fileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            errorForInvalidUrl = string.Empty;
            Uri fileLocation;
            if (!Uri.TryCreate(fileUrl, UriKind.Absolute, out fileLocation))
            {
                errorForInvalidUrl = string.Format(@"The file URL should be a valid absolute URL. Current:[{0}]", fileUrl);
                return false;
            }

            string fileName = Path.GetFileName(fileUrl);
            if (string.IsNullOrEmpty(fileName))
            {   
                errorForInvalidUrl = string.Format(@"The file URL should point to a file. Current:[{0}]", fileUrl);
                return false;
            }
            
            // If all validation are passed, return true.
            return true;
        }

        /// <summary>
        /// A method used to check a file URL whether is a valid file URL. If the URL is not a valid file URL, this method will raise a UriFormatException.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the file URL which will be validated.</param>
        private static void VerifyFileUrl(string fileUrl)
        { 
            string errorMsg;
            if (!TryVerifyFileUrl(fileUrl, out errorMsg))
            {
                throw new UriFormatException(errorMsg);
            }
        }
 
        /// <summary>
        /// A method used to get unique resource name by specified value.
        /// </summary>
        /// <param name="resourceName">A parameter represents the resource name which is used to construct the unique resource name.</param>
        /// <param name="isMultipleResourcesPerCase">A parameter represents bool value indicating this method is used for scenario "multiple resources per case". The default value is false.</param>
        /// <returns>A return value represents the generated unique resource name.</returns>
        private string GetUniqueResourceName(string resourceName, bool isMultipleResourcesPerCase = false)
        {
            if (string.IsNullOrEmpty(resourceName))
            {
                throw new ArgumentNullException("resourceName");
            }

            string currentResourceTemp;
            if (isMultipleResourcesPerCase)
            {
                FileNameCounterOfPerTestCases += 1;
                currentResourceTemp = Common.GenerateResourceName(this.Site, resourceName, FileNameCounterOfPerTestCases) + ".txt";
            }
            else
            {
                currentResourceTemp = Common.GenerateResourceName(this.Site, resourceName) + ".txt";
            }

            return currentResourceTemp;
        }

        #endregion 
    }
}