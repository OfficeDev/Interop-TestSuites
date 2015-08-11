namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite base class of MS-COPYS.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region variables
        /// <summary>
        /// Gets or sets an instance of IMS_COPYSAdapter
        /// </summary>
        protected static IMS_COPYSAdapter MSCopysAdapter { get; set; }
        
        /// <summary>
        /// Gets or sets an instance of IMS_COPYSSutControlAdapter
        /// </summary>
        protected static IMS_COPYSSUTControlAdapter MSCOPYSSutControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets an uint indicate the file name number value on current test case.
        /// </summary>
        protected static uint FileNameCounterOfPerTestCases { get; set; }

        /// <summary>
        ///  Gets or sets an base64 string indicates the file contents of source file.
        /// </summary>
        protected static string SourceFileContentBase64Value { get; set; }

        /// <summary>
        /// Gets or sets a list type instance used to record all URLs of files added by TestSuiteHelper
        /// </summary>
        private static List<string> FilesUrlRecordOfDestination { get; set; }
        #endregion variables

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            // A method is used to initialize the variables.
            TestClassBase.Initialize(testContext);
            if (null == MSCopysAdapter)
            {
                MSCopysAdapter = BaseTestSite.GetAdapter<IMS_COPYSAdapter>();
            }

            if (null == MSCOPYSSutControlAdapter)
            {
                MSCOPYSSutControlAdapter = BaseTestSite.GetAdapter<IMS_COPYSSUTControlAdapter>();
            }

            if (null == FilesUrlRecordOfDestination)
            {
                FilesUrlRecordOfDestination = new List<string>();
            }

            if (string.IsNullOrEmpty(SourceFileContentBase64Value))
            {
                string sourceFileContent = Common.GetConfigurationPropertyValue("SourceFileContents", BaseTestSite);
                byte[] souceFileContentBinaries = Encoding.UTF8.GetBytes(sourceFileContent);
                SourceFileContentBase64Value = Convert.ToBase64String(souceFileContentBinaries);
            }
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            if (null != FilesUrlRecordOfDestination && 0 != FilesUrlRecordOfDestination.Count)
            {
                StringBuilder strBuilder = new StringBuilder();
                foreach (string fileUrlItem in FilesUrlRecordOfDestination)
                {
                    strBuilder.Append(fileUrlItem + @";");
                }

                string fileUrlValues = strBuilder.ToString();

                bool isDeleteAllFilesSuccessful;

                isDeleteAllFilesSuccessful = MSCOPYSSutControlAdapter.DeleteFiles(fileUrlValues);

                if (!isDeleteAllFilesSuccessful)
                {
                    throw new InvalidOperationException("Not all the files are deleted clearly.");
                }
            }

            TestClassBase.Cleanup();
        }

        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        { 
            Common.CheckCommonProperties(this.Site, true);

            // Reset to the default user.
            MSCopysAdapter.SwitchUser(TestSuiteManageHelper.DefaultUser, TestSuiteManageHelper.PasswordOfDefaultUser, TestSuiteManageHelper.DomainOfDefaultUser);

            // Initialize the unique resource counter
            FileNameCounterOfPerTestCases = 0; 
        }
 
        #endregion Test Suite Initialization

        #region protected methods

        /// <summary>
        /// A method is used to generate the destination URL. The file name in the URL is unique.
        /// </summary>
        /// <param name="destinationType">A parameter represents the destination URL type.</param>
        /// <returns>A return value represents the destination URL.</returns>
        protected string GetDestinationFileUrl(DestinationFileUrlType destinationType)
        {
            string urlPatternValueOfDestinationFile = string.Empty;
            string expectedPropertyName = string.Empty;
            switch (destinationType)
            {
                case DestinationFileUrlType.NormalDesLibraryOnDesSUT:
                    {
                        expectedPropertyName = "UrlPatternOfDesFileOnDestinationSUT";
                        break;
                    }

                case DestinationFileUrlType.MWSLibraryOnDestinationSUT:
                    {
                        expectedPropertyName = "UrlPatternOfDesFileForMWSOnDestinationSUT";
                        break;
                    }

                default:
                    {
                        throw new InvalidOperationException("The test suite only supports two destination type: [NormalDesLibraryOnDesSUT] and [MWSLibraryOnDestinationSUT].");
                    }
            }

            // Get match URL pattern.
            urlPatternValueOfDestinationFile = Common.GetConfigurationPropertyValue(expectedPropertyName, this.Site);

            if (urlPatternValueOfDestinationFile.IndexOf("{FileName}", StringComparison.OrdinalIgnoreCase) <= 0)
            {
                throw new InvalidOperationException(string.Format(@"The [{0}] property should contain the ""{fileName}"" placeholder.", expectedPropertyName));
            }

            string fileNameValue = this.GetUniqueFileName();
            string actualDestinationFileUrl = urlPatternValueOfDestinationFile.ToLower().Replace("{filename}", fileNameValue);

            // Verify the URL whether point to a file.
            FileUrlHelper.ValidateFileUrl(actualDestinationFileUrl);

            return actualDestinationFileUrl;
        }

        /// <summary>
        /// A method used to get source file URL. The URL's file name is unique.
        /// </summary>
        /// <param name="sourceFileType">A parameter represents the source URL type.</param>
        /// <returns>A return value represents the destination URL.</returns>
        protected string GetSourceFileUrl(SourceFileUrlType sourceFileType)
        {
            string expectedPropertyName;
            switch (sourceFileType)
            {
                case SourceFileUrlType.SourceFileOnDesSUT:
                    {
                        expectedPropertyName = "SourceFileUrlOnDesSUT";
                        break;
                    }

                case SourceFileUrlType.SourceFileOnSourceSUT:
                    {
                        expectedPropertyName = "SourceFileUrlOnSourceSUT";
                        break;
                    }

                default:
                    {
                        throw new InvalidOperationException("The test suite only support two source URL type: [SourceFileUrlOnDesSUT] and [SourceFileUrlOnSourceSUT].");
                    }
            }

            string expectedSourceFileUrl = Common.GetConfigurationPropertyValue(expectedPropertyName, this.Site);

            // Verify the URL whether point to a file.
            FileUrlHelper.ValidateFileUrl(expectedSourceFileUrl);

            return expectedSourceFileUrl;
        }

        /// <summary>
        /// A method is used to get a unique name of file which is used to upload into document library list.
        /// </summary>
        /// <returns>A return value represents the unique name that is combined with the file Object name and time stamp</returns>
        protected string GetUniqueFileName()
        {
            FileNameCounterOfPerTestCases += 1;

            // Sleep 1 seconds in order to ensure "Common.GenerateResourceName" method will not generate same file names for invoke it twice. 
            Thread.Sleep(1000);
            string fileName = Common.GenerateResourceName(this.Site, "file", FileNameCounterOfPerTestCases);
            return string.Format("{0}.txt", fileName);
        }

        /// <summary>
        /// A method used to select a field by specified field attribute value. The "DisplayName" is not a unique identifier for a field, so the method return the first match item when using "DisplayName" attribute value. If there are no any fields are found, this method will raise an InvalidOperationException.
        /// </summary>
        /// <param name="fields">A parameter represents the fields array.</param>
        /// <param name="expectedAttributeValue">A parameter represents the expected attribute value which is used to match field item.</param>
        /// <param name="usedAttribute">A parameter represents the attribute type which is used to compare with the expected attribute value.</param>
        /// <returns>A return value represents the selected field.</returns>
        protected FieldInformation SelectFieldBySpecifiedAtrribute(FieldInformation[] fields, string expectedAttributeValue, FieldAttributeType usedAttribute)
        {
            FieldInformation selectedField = this.FindOutTheField(fields, expectedAttributeValue, usedAttribute);
            if (null == selectedField)
            {
               string errorMsg = string.Format(
                        "Could not find the expected field by specified value:[{0}] of [{1}] attribute.",
                        expectedAttributeValue,
                        usedAttribute);

               throw new InvalidOperationException(errorMsg);
            }

            return selectedField;
        }

        /// <summary>
        /// A method used to verify whether all copy results are successful. If there is any non-success copy result, it raise an exception.
        /// </summary>
        /// <param name="copyResults">A parameter represents the copy results collection.</param>
        /// <param name="raiseAssertEx">A parameter indicates the method whether raise an assert exception, if not all the copy result are successful. The 'true' means raise an exception.</param>
        /// <returns>Return 'true' indicating all copy results are success, otherwise return 'false'. If the "raiseAssertEx" is set to true, the method will raise a exception instead of returning 'false'.</returns>
        protected bool VerifyAllCopyResultsSuccess(CopyResult[] copyResults, bool raiseAssertEx = true)
        {
           if (null == copyResults)
           {
               throw new ArgumentNullException("copyResults");
           }

           if (0 == copyResults.Length)
           {
               throw new ArgumentException("The copy results collection should contain at least one item", "copyResults");
           }

           var notSuccessResultItem = from resultItem in copyResults
                                     where CopyErrorCode.Success != resultItem.ErrorCode
                                     select resultItem;

           bool verifyResult = false;
           if (notSuccessResultItem.Count() > 0)
           {
               StringBuilder strBuilderOfErrorResult = new StringBuilder();
               foreach (CopyResult resultItem in notSuccessResultItem)
               {   
                   string errorMsgItem = string.Format(
                               "ErrorCode:[{0}]\r\nErrorMessage:[{1}]\r\nDestinationUrl:[{2}]",
                               resultItem.ErrorCode,
                               string.IsNullOrEmpty(resultItem.ErrorMessage) ? "None" : resultItem.ErrorMessage,
                               resultItem.DestinationUrl);

                   strBuilderOfErrorResult.AppendLine(errorMsgItem);
                   strBuilderOfErrorResult.AppendLine("===========");
               }

               this.Site.Log.Add(
                               LogEntryKind.Debug,
                               "Total non-successful copy results:[{0}]\r\n{1}",
                               notSuccessResultItem.Count(),
                               strBuilderOfErrorResult.ToString());

               if (raiseAssertEx)
               {
                  this.Site.Assert.Fail("Not all copy result are successful.");
               }
           }
           else
           {
               verifyResult = true;
           }

           return verifyResult;
        }

        /// <summary>
        /// A method used to generate invalid file URL by confusing the folder path. The method will confuse the path part which is nearest to the file name part.
        /// </summary>
        /// <param name="originalFilePath">A parameter represents the original file path.</param>
        /// <returns>A return value represents a file URL with an invalid folder path.</returns>
        protected string GenerateInvalidFolderPathForFileUrl(string originalFilePath)
        {
            string fileName = FileUrlHelper.ValidateFileUrl(originalFilePath);
            string directoryName = Path.GetDirectoryName(originalFilePath);

            // Append a guid value to ensure the folder name is not a valid folder.
            string invalidDirectoryName = directoryName + Guid.NewGuid().ToString("N");
            string invalidFileUrl = Path.Combine(invalidDirectoryName, fileName);

            // Work around for local path format mapping to URL path format.
            invalidFileUrl = invalidFileUrl.Replace(@"\", @"/");
            invalidFileUrl = invalidFileUrl.Replace(@":/", @"://");
            return invalidFileUrl;
        }

        /// <summary>
        /// A method used to generate invalid file URL by construct a not-existing file name. 
        /// </summary>
        /// <param name="originalFilePath">A parameter represents the original file path, it must be a valid URL.</param>
        /// <returns>A return value represents a file URL with a not-existing file name.</returns>
        protected string GenerateInvalidFileUrl(string originalFilePath)
        {
            FileUrlHelper.ValidateFileUrl(originalFilePath);
            string directoryName = Path.GetDirectoryName(originalFilePath);

            // Append an invalid URL char to file name is not.
            string invalidFileName = string.Format(@"Invalid{0}File{1}Name.txt", @"%", @"&");
            string invalidFileUrl = Path.Combine(directoryName, invalidFileName);

            // Work around for local path format mapping to URL path format.
            invalidFileUrl = invalidFileUrl.Replace(@"\", @"/");
            invalidFileUrl = invalidFileUrl.Replace(@":/", @"://");
            return invalidFileUrl;
        }
        
        /// <summary>
        /// A method used to collect a file by specified file URL. The test suite will try to delete all collect files.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the URL of a file which will be collected to delete. The file must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        protected void CollectFileByUrl(string fileUrl)
        {
           FileUrlHelper.ValidateFileUrl(fileUrl);
           this.CollectFileToRecorder(fileUrl);
        }

        /// <summary>
        /// A method used to collect files from specified file URLs. The test suite will try to delete all collect files.
        /// </summary>
        /// <param name="fileUrls">A parameter represents the arrays of file URLs of files which will be collected to delete. All the files must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        protected void CollectFileByUrl(string[] fileUrls)
        {
            if (null == fileUrls || 0 == fileUrls.Length)
            {
                throw new ArgumentNullException("fileUrls");
            }

            foreach (string fileUrlItem in fileUrls)
            {
                this.CollectFileByUrl(fileUrlItem);
            }
        }

        /// <summary>
        /// A method used to upload a txt file to the destination SUT.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the expected URL of a file which will be uploaded into the destination SUT. The file URL must point to the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        protected void UploadTxtFileByFileUrl(string fileUrl)
        {
            if (null == MSCOPYSSutControlAdapter)
            {
                throw new InvalidOperationException("The SUT control adapter is not initialized.");
            }

            FileUrlHelper.ValidateFileUrl(fileUrl);
            bool isFileUploadSuccessful = MSCOPYSSutControlAdapter.UploadTextFile(fileUrl);
            if (!isFileUploadSuccessful)
            {
                this.Site.Assert.Fail("Could not upload a txt file to the destination SUT. Expected file URL:[{0}]", fileUrl);
            }

            this.CollectFileToRecorder(fileUrl);
        }

        #endregion protected methods

        #region private methods

        /// <summary>
        /// A method used to collect a file by specified file URL.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the URL of a file which will be collected to delete.</param>
        private void CollectFileToRecorder(string fileUrl)
        {
            if (null == FilesUrlRecordOfDestination)
            {
                FilesUrlRecordOfDestination = new List<string>();
            }

            FilesUrlRecordOfDestination.Add(fileUrl);
        }

        /// <summary>
        /// A method used to find out a field by specified field attribute value. The "DisplayName" is not a unique identifier for a field, so the method return the first match item when using "DisplayName" attribute value.
        /// </summary>
        /// <param name="fields">A parameter represents the fields array.</param>
        /// <param name="expectedAttributeValue">A parameter represents the expected attribute value which is used to match field item.</param>
        /// <param name="usedAttribute">A parameter represents the attribute type which is used to compare with the expected attribute value.</param>
        /// <returns>A return value represents the selected field.</returns>
        private FieldInformation FindOutTheField(FieldInformation[] fields, string expectedAttributeValue, FieldAttributeType usedAttribute)
        {
            #region validate parameters
            if (null == fields)
            {
                throw new ArgumentNullException("fields");
            }

            if (string.IsNullOrEmpty(expectedAttributeValue))
            {
                throw new ArgumentNullException("attributeValue");
            }

            if (0 == fields.Length)
            {
                throw new ArgumentException("The fields' array should contain at least one item.");
            }

            #endregion validate parameters

            FieldInformation selectedField = null;

            #region select the field by specified attribute value
            switch (usedAttribute)
            {
                case FieldAttributeType.InternalName:
                    {
                        var selectResult = from fieldItem in fields
                                           where expectedAttributeValue.Equals(fieldItem.InternalName, StringComparison.OrdinalIgnoreCase)
                                           select fieldItem;

                        selectedField = 1 == selectResult.Count() ? selectResult.ElementAt<FieldInformation>(0) : null;
                        break;
                    }

                case FieldAttributeType.Id:
                    {
                        Guid expectedGuidValue;
                        if (!Guid.TryParse(expectedAttributeValue, out expectedGuidValue))
                        {
                            throw new InvalidOperationException(
                                        string.Format(
                                                       @"The attributeValue parameter value should be a valid GUID value when select field by using [Id] attribute. Current attributeValue parameter:[{0}].",
                                                       expectedAttributeValue));
                        }

                        var selectResult = from fieldItem in fields
                                           where expectedGuidValue.Equals(fieldItem.Id)
                                           select fieldItem;

                        selectedField = 1 == selectResult.Count() ? selectResult.ElementAt<FieldInformation>(0) : null;
                        break;
                    }

                case FieldAttributeType.DisplayName:
                    {
                        var selectResult = from fieldItem in fields
                                           where expectedAttributeValue.Equals(fieldItem.InternalName, StringComparison.OrdinalIgnoreCase)
                                           select fieldItem;

                        // The DisplayName attribute is not unique.
                        selectedField = selectResult.Count() > 0 ? selectResult.ElementAt<FieldInformation>(0) : null;
                        break;
                    }
            }
            #endregion select the field by specified attribute value

            if (null == selectedField)
            {
                string errorMsg = string.Format(
                         "Could not find the expected field by specified vale:[{0}] of [{1}] attribute.",
                         expectedAttributeValue,
                         usedAttribute);

                #region log the fields information

                StringBuilder logBuilder = new StringBuilder();
                logBuilder.AppendLine(errorMsg);
                logBuilder.AppendLine("Fields in current field collection:");
                foreach (FieldInformation fieldItem in fields)
                {
                    string fieldInformation = string.Format(
                                @"FieldName:[{0}] InternalName:[{1}] Type:[{2}] Value:[{3}]",
                                fieldItem.DisplayName,
                                string.IsNullOrEmpty(fieldItem.InternalName) ? "Null/Empty" : fieldItem.InternalName,
                                fieldItem.Type,
                                string.IsNullOrEmpty(fieldItem.Value) ? "Null/Empty" : fieldItem.Value);

                    logBuilder.AppendLine(fieldInformation);
                }

                logBuilder.AppendLine("==================");
                logBuilder.AppendLine(string.Format("Total fields:[{0}]", fields.Length));
                this.Site.Log.Add(LogEntryKind.Debug, logBuilder.ToString());
                #endregion log the fields information
            }

            return selectedField;
        }
        #endregion private methods
    }
}