namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with valid or invalid parameters.
    /// <list type="bullet">
    ///     <item>CheckInFile</item>
    ///     <item>CheckOutFile</item>
    ///     <item>UndoCheckOut</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S05_OperationOnFiles : TestClassBase
    {
        #region Private member variables

        /// <summary>
        /// Protocol adapter
        /// </summary>
        private IMS_LISTSWSAdapter listwsInstance;

        /// <summary>
        /// SUTControl adapter used to upload file into document library.
        /// </summary>
        private IMS_LISTSWSSUTControlAdapter sutControlAdapter;

        #endregion

        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The text context</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases

        #region CheckInFile

        /// <summary>
        /// This test case is used to test that SOAP fault returns when the input pageUrl parameter does not refer to a document library in CheckInFile operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC01_CheckInFile_WithoutDocument()
        {
            // create a normal document library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Invoke CheckInFile operation with pageUrl not referring to a document library, and catch the exception.

            string errorCode = string.Empty;
            bool isSoapFaultExisted = false;
            try
            {
                this.listwsInstance.CheckInFile(absoluteFileUrl + TestSuiteHelper.GenerateRandomString(1), string.Empty, "0");
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture R393 R1642
            // If Soap Fault Existed, Capture R393
            Site.CaptureRequirementIfIsTrue(
                                 isSoapFaultExisted,
                                 393,
                                 @"[In CheckInFile operation] If the pageUrl does not refer to a document, the protocol server MUST return a SOAP fault.");

            // If there is no any error code in soap fault, capture R1642
            Site.CaptureRequirementIfIsTrue(
                        string.IsNullOrEmpty(errorCode),
                        1642,
                        "[In CheckInFile operation] [If the pageUrl does not refer to a document library, the protocol server MUST return a SOAP fault.]There is no error code returned for this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify that the server returns a SOAP fault without error code when the checkInType parameter is an empty string in CheckInFile operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC02_CheckInFile_EmptyCheckInType()
        {
            // create a normal list document Library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listTitle, (int)TemplateType.Document_Library);
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            // Check out the added file.
            bool isSoapFautExisted = false;
            isSoapFautExisted = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            Site.Assert.IsTrue(isSoapFautExisted, "CheckOutFile must succeed!");

            // Check in file with null "comment" parameter and valid "checkinType" parameter.
            isSoapFautExisted = this.listwsInstance.CheckInFile(absoluteFileUrl, null, CheckInTypeValue.MajorCheckIn);
            Site.Assert.IsTrue(isSoapFautExisted, "CheckInFile must succeed!");

            // Check out the added file again.
            isSoapFautExisted = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            Site.Assert.IsTrue(isSoapFautExisted, "CheckOutFile must succeed!");

            // Invoke CheckInFile operation with pageUrl not referring to a document library, and catch the exception.
            string errorCode = string.Empty;
            isSoapFautExisted = false;
            try
            {
                string emptyCheckInTypeValue = string.Empty;
                this.listwsInstance.CheckInFile(absoluteFileUrl, null, emptyCheckInTypeValue);
            }
            catch (SoapException soapEx)
            {
                isSoapFautExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // If there is a soap fault  capture R395
            Site.CaptureRequirementIfIsTrue(
                                        isSoapFautExisted,
                                        395,
                                        @"[In CheckInFile operation] If the CheckinType element is an empty string, the protocol server MUST return a SOAP fault.");

            // If there is a soap fault and no error code,  capture R1643
            Site.CaptureRequirementIfIsTrue(
                                       isSoapFautExisted && string.IsNullOrEmpty(errorCode),
                                       1643,
                                       @"[In CheckInFile operation] [If the checkInType parameter is an empty string, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");
        }

        /// <summary>
        /// This test case is used to test CheckInFile operation when at least one of its input parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC03_CheckInFile_InvalidParameter()
        {
            // create a normal document Library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file.

            bool isFileCheckOutSuccessfully = false;
            isFileCheckOutSuccessfully = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);

            Site.Assert.IsTrue(isFileCheckOutSuccessfully, "CheckOutFile must succeed!");

            #endregion

            #region Try to check in the file which has been checked out with invalid pageUrl parameter.

            bool isSoapFaultExisted = false;
            string errorCode = string.Empty;
            try
            {
                string nullPageUrl = null;
                this.listwsInstance.CheckInFile(nullPageUrl, string.Empty, CheckInTypeValue.MajorCheckIn);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            #endregion

            #region Capture R390 if the error code 0x82000001 is returned.

            // if there are a soap fault and the error code equal to the "0x82000001", capture R390
            Site.CaptureRequirementIfIsTrue(
                                isSoapFaultExisted && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                390,
                                @"[In CheckInFile operation] If the pageUrl is null the protocol server MUST return a SOAP fault with error code 0x82000001.");
            #endregion

            #region Try to check in the file if the pageUrl is empty string.

            isSoapFaultExisted = false;
            errorCode = string.Empty;
            try
            {
                string emptyPageUrl = string.Empty;
                this.listwsInstance.CheckInFile(emptyPageUrl, string.Empty, CheckInTypeValue.MajorCheckIn);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // if there are a soap fault and the error code equal to the "0x82000001", capture R391
            Site.CaptureRequirementIfIsTrue(
                                    isSoapFaultExisted && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                    391,
                                    @"[In CheckInFile operation] If the pageUrl is empty string the protocol server MUST return a SOAP fault with error code 0x82000001.");

            // if there are a soap fault and the error code equal to the "0x82000001", capture 1640 
            Site.CaptureRequirementIfIsTrue(
                                 isSoapFaultExisted && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                1640,
                                @"[In CheckInFile operation] [If the pageUrl is empty string the protocol server MUST return a SOAP fault with error code 0x82000001.] This indicates that the parameter pageUrl is missing.");

            #endregion

            #region Try to check in the file if pageUrl setting to an invalid URL.

            isSoapFaultExisted = false;
            string errorString = string.Empty;
            errorCode = string.Empty;
            try
            {
                string invalidPageUrl = "/";
                this.listwsInstance.CheckInFile(invalidPageUrl, string.Empty, CheckInTypeValue.MajorCheckIn);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorString = TestSuiteHelper.GetErrorString(soapEx);
            }

            if (Common.IsRequirementEnabled(3921, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R3921. SOAP fault {0}.", isSoapFaultExisted ? "is returned with error string '" + errorString + "'" : "is not returned");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R3921
                bool isVerifiedR3921 = isSoapFaultExisted && string.Equals(errorString, "Invalid URI: The format of the URI could not be determined.");

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR3921,
                    3921,
                    @"[In CheckInFile operation] [If the pageUrl is an invalid URL] Implementation does return a SOAP fault with error string ""Invalid URI: The format of the URI could not be determined"".(SharePoint Foundation 2010 and above follow this behavior.)");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1641
                this.Site.CaptureRequirementIfIsTrue(
                                         isVerifiedR3921,
                                         1641,
                                         @"[In CheckInFile operation] [If the pageUrl is an invalid URL, the protocol server SHOULD<41> return a SOAP fault with error string ""Invalid URI: The format of the URI could not be determined"".] This indicates that the parameter pageUrl is invalid.");
            }

            if (Common.IsRequirementEnabled(3922, this.Site))
            {
                this.Site.CaptureRequirementIfIsFalse(
                    isSoapFaultExisted,
                    3922,
                    @"[In CheckInFile operation] [If the pageUrl is an invalid URL] Implementation does not return a SOAP fault.(<41> Section 3.1.4.7:  Windows SharePoint Services 3.0 does not return a SOAP fault..)");
            }

            #endregion

            #region Try to check in the file that have been check in success

            bool isCheckInSuccess = false;

            // Call the "CheckInFile" operation and ensure check in successfully.  
            isCheckInSuccess = this.listwsInstance.CheckInFile(absoluteFileUrl, string.Empty, CheckInTypeValue.MajorCheckIn);
            Site.Assert.IsTrue(isCheckInSuccess, "CheckInFile operation must be successfully for the first time");

            // Call the "CheckInFile" operation to check in the file which have been check in successfully in previous checked in file operation.
            isCheckInSuccess = this.listwsInstance.CheckInFile(absoluteFileUrl, string.Empty, CheckInTypeValue.MajorCheckIn);

            // If "CheckInFile" operation return false, capture R2304
            Site.CaptureRequirementIfIsFalse(
                isCheckInSuccess,
                2304,
                @"[CheckInFileResponse][CheckInFileResult] [The value is True if the operation is successful;] otherwise, False is returned.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the CheckInFile operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC04_CheckInFile_Succeed()
        {
            // create a list by using document library template
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file.
            bool checkOutSucceeded = false;
            checkOutSucceeded = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            Site.Assert.IsTrue(checkOutSucceeded, "CheckOutFile must succeed!");
            #endregion

            #region Check in the file which has been checked out.
            bool checkInSucceeded = false;

            // generate a random string.
            string checkInComments = TestSuiteHelper.GenerateRandomString(5);
            checkInSucceeded = this.listwsInstance.CheckInFile(absoluteFileUrl, checkInComments, CheckInTypeValue.MajorCheckIn);

            #endregion

            #region Capture R16381, R398 and R1660 if the CheckInFile succeeds and returns true value.

            // Verify requirement R16381.
            // If there are no other errors, it means implementation does support this CheckInFile method. R16381 can be captured.
            if (Common.IsRequirementEnabled(16381, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                    checkInSucceeded,
                    16381,
                    @"Implementation does support this method[CheckInFile]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If CheckInFile operation return true, capture R398, 1660
            Site.CaptureRequirementIfIsTrue(
                checkInSucceeded,
                398,
                @"[In CheckInFile operation] If there are no other errors, the document located at pageUrl MUST be checked-in by using comments and CheckinType specified in the CheckInFileSoapIn request message.");

            Site.CaptureRequirementIfIsTrue(
                checkInSucceeded,
                1660,
                @"[CheckInFileResponse][CheckInFileResult]The value is True if the operation is successful;");
            #endregion
        }

        #endregion

        #region CheckOutFile

        /// <summary>
        /// This test case is used to test that the SOAP fault returns when the input pageUrl parameter does not refer to a document library in CheckOutFile operation.
        /// </summary>  
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC05_CheckOutFile_WithInvalidDocument()
        {
            // create a normal document Library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Invoke CheckOutFile operation with pageUrl not referring to a document library, and catch the exception.

            string errorCode = string.Empty;
            bool isSoapFaultExisted = false;
            try
            {
                this.listwsInstance.CheckOutFile(absoluteFileUrl + TestSuiteHelper.GenerateRandomString(2), bool.FalseString, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture R409 R1665
            // If Soap Fault Existed, Capture R409
            Site.CaptureRequirementIfIsTrue(
                                    isSoapFaultExisted,
                                    409,
                                    @"[In CheckOutFile operation]If the pageUrl does not refer to a document, the protocol server MUST return a SOAP fault.");

            // If there is no any error code in soap fault, capture R1665
            Site.CaptureRequirementIfIsTrue(
                                     string.IsNullOrEmpty(errorCode),
                                     1665,
                                     @"[In CheckOutFile operation] [If the pageUrl does not refer to a document, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the CheckOutFile operation when at least one of its input 
        /// parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC06_CheckOutFile_InvalidParameter()
        {
            // create a normal document Library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file with null pageUrl parameter, try to capture R406; R2305;

            bool isSoapFaultExist = false;
            string errorCode = string.Empty;
            try
            {
                string nullPageUrl = null;
                this.listwsInstance.CheckOutFile(nullPageUrl, bool.TrueString, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // if there are a soap fault and the error code equal to the "0x82000001", capture R406
            Site.CaptureRequirementIfIsTrue(
                                isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                406,
                                @"[In CheckOutFile operation]If the pageUrl is null  the protocol server MUST return a SOAP fault with error code 0x82000001.");

            #endregion

            #region Check out the added file with pageUrl is empty string, and try to capture R407; R1663

            isSoapFaultExist = false;
            errorCode = string.Empty;
            try
            {
                string emptyPageUrl = string.Empty;
                this.listwsInstance.CheckOutFile(emptyPageUrl, bool.TrueString, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // If there are a soap fault and the error code equal to the "0x82000001", capture R407
            Site.CaptureRequirementIfIsTrue(
                                 isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                 407,
                                 @"[In CheckOutFile operation] If the pageUrl is empty string, the protocol server MUST return a SOAP fault with error code 0x82000001.");

            // If there are a soap fault and the error code equal to the "0x82000001", capture R1663
            Site.CaptureRequirementIfIsTrue(
                                isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                1663,
                                @"[In CheckOutFile operation] [If the pageUrl is empty string, the protocol server MUST return a SOAP fault with error code 0x82000001.] This indicates that the pageUrl is an empty string.");
            #endregion

            #region Check out the added file with pageUrl setting to an invalid URL, try to capture MS-LISTSWS_R4081 MS-LISTSWS_R1664 and MS-LISTSWS_R4082

            isSoapFaultExist = false;
            string errorString = string.Empty;
            errorCode = string.Empty;
            try
            {
                string invalidPageUrl = "/";
                this.listwsInstance.CheckOutFile(invalidPageUrl, bool.TrueString, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorString = TestSuiteHelper.GetErrorString(soapEx);
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            if (Common.IsRequirementEnabled(4081, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R4081. SOAP fault {0}.", isSoapFaultExist ? "is returned with error string '" + errorString + "'" : "is not returned");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R4081
                bool isVerifiedR4081 = isSoapFaultExist && string.Equals(errorString, "Invalid URI: The format of the URI could not be determined.");

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR4081,
                    4081,
                    @"[In CheckOutFile operation] [If the pageUrl is an invalid URL] Implementation does return a SOAP fault with error string ""Invalid URI: The format of the URI could not be determined."".(SharePoint Foundation 2010 and above follow this behavior.)");

                // If there is no any error code in soap fault, capture MS-LISTSWS_R1664
                this.Site.CaptureRequirementIfIsTrue(
                                         string.IsNullOrEmpty(errorCode),
                                         1664,
                                         @"[In CheckOutFile operation] [If the pageUrl is an invalid URL, the protocol server SHOULD<38> return a SOAP fault with error string ""Invalid URI: The format of the URI could not be determined."". ]There is no error code for this fault.");
            }

            if (Common.IsRequirementEnabled(4082, this.Site))
            {
                this.Site.CaptureRequirementIfIsFalse(
                    isSoapFaultExist,
                    4082,
                    @"[In CheckOutFile operation] [If the pageUrl is an invalid URL] Implementation does not return a SOAP fault.(<43> Section 3.1.4.8:  wss3 does not return a SOAP fault.)");
            }

            #endregion

            #region Check out the added file with the checkoutToLocal parameter does not resolve to a valid Boolean string.Try to capture R410;R1666;

            isSoapFaultExist = false;
            errorCode = string.Empty;
            try
            {
                // Get an invalid path
                string invalidBoolValue = Guid.NewGuid().ToString("N");
                this.listwsInstance.CheckOutFile(absoluteFileUrl, invalidBoolValue, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // If there are a Soap fault, capture R410
            Site.CaptureRequirementIfIsTrue(
                               isSoapFaultExist,
                               410,
                               @"[In CheckOutFile operation] If the checkoutToLocal parameter does not resolve to a valid Boolean string (case-insensitive equality to ""True"" or ""False"", ignoring leading and trailing white space), the protocol server MUST return a SOAP fault.");

            // If there are a soap fault and no error code return in SoapFault, capture R1666
            Site.CaptureRequirementIfIsTrue(
                                isSoapFaultExist && string.IsNullOrEmpty(errorCode),
                                1666,
                                @"[In CheckOutFile operation] [If the checkoutToLocal parameter does not resolve to a valid Boolean string (case-insensitive equality to ""True"" or ""False"", ignoring leading and trailing white space), the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            #endregion

            isSoapFaultExist = false;
            errorCode = string.Empty;
            try
            {
                // Get an invalid path
                string invalidBoolValue = Guid.NewGuid().ToString("N");
                this.listwsInstance.CheckOutFile(absoluteFileUrl, invalidBoolValue, string.Empty);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            #region Check out a file that have been checked out, if the CheckOutFile should return false, capture R2305;

            bool ischeckOutSuccess = false;
            ischeckOutSuccess = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            Site.Assert.IsTrue(ischeckOutSuccess, "The CheckOutFile operation must be successful in the first time.");

            // Check out a file that have been checked out
            ischeckOutSuccess = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);

            // if the second CheckOutFile operation return false, capture R2305
            Site.CaptureRequirementIfIsFalse(
               ischeckOutSuccess,
               2305,
               @"[CheckOutFileResponse][The value is True if the operation is successful; ]otherwise, False is returned");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the CheckOutFile operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC07_CheckOutFile_Succeed()
        {
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file with all valid parameters.

            bool isCheckout = false;

            isCheckout = this.listwsInstance.CheckOutFile(absoluteFileUrl, "False", null);
            #endregion

            #region Capture R16611, R412 and R1679 if the CheckOutFile succeeds and returns true value.

            // Verify requirement R16611.
            // If there are no other errors, it means implementation does support this CheckOutFile method. R16611 can be captured.
            if (Common.IsRequirementEnabled(16611, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                    isCheckout,
                    16611,
                    @"Implementation does support this method[CheckOutFile]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            Site.CaptureRequirementIfIsTrue(
                isCheckout,
                412,
                @"[In CheckOutFile operation] If there are no other errors, the document located at pageUrl MUST be checked out by using checkoutToLocal and last modified as specified in the CheckOutFileSoapIn request message.");

            Site.CaptureRequirementIfIsTrue(
                isCheckout,
                1679,
                @"[CheckOutFileResponse]The value is True if the operation is successful;");
            #endregion
        }

        #endregion

        #region UndoCheckOut

        /// <summary>
        /// This test case is used to test that the SOAP fault returns when the input pageUrl parameter does not refer to a document library in UndoCheckOut operation.
        /// </summary>  
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC08_UndoCheckOut_WithNoDocument()
        {
            // create a normal document Library and upload a file to SUT
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Invoke UndoCheckOut operation with pageUrl not referring to a document library, and catch the exception.

            string errorCode = string.Empty;
            bool isSoapFaultExisted = false;
            try
            {
                this.listwsInstance.UndoCheckOut(absoluteFileUrl + TestSuiteHelper.GenerateRandomString(2));
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture R786 R1964
            // If Soap Fault Existed, Capture R786
            Site.CaptureRequirementIfIsTrue(
                              isSoapFaultExisted,
                              786,
                              "[In UndoCheckOut operation] If the pageUrl does not refer to a document, the protocol server MUST return a SOAP fault. ");

            // If there is no any error code in soap fault, capture R1964
            Site.CaptureRequirementIfIsTrue(
                               string.IsNullOrEmpty(errorCode),
                               1964,
                               "[In UndoCheckOut operation] [If the pageUrl does not refer to a document, the protocol server MUST return a SOAP fault.] There is no error code for this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the UndoCheckOut operation when at least one of its input parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC09_UndoCheckOut_InvalidParameter()
        {
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file.

            bool checkOutSucceeded = false;
            checkOutSucceeded = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            if (!checkOutSucceeded)
            {
                Site.Assert.Fail("Expect CheckOutFile operation failed due to unexpected reason.");
            }

            #endregion

            #region Undo the checkout with null pageUrl parameter, try to capture R783 and R2306

            bool isSoapFaultExist = false;
            string errorCode = string.Empty;
            try
            {
                string nullPageUrl = null;
                this.listwsInstance.UndoCheckOut(nullPageUrl);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R783 if the error code 0x82000001 is returned.
            Site.CaptureRequirementIfIsTrue(
                                        isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                        783,
                                        @"[In UndoCheckOut operation] If the pageUrl is null the protocol server MUST return a SOAP fault with error code 0x82000001.");

            #endregion

            #region Undo the checkout with pageUrl parameter setting to an invalid URL, try to capture R7851 and R7852

            isSoapFaultExist = false;
            string errorString = string.Empty;
            try
            {
                string invalidPageUrl = "/";
                this.listwsInstance.UndoCheckOut(invalidPageUrl);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorString = TestSuiteHelper.GetErrorString(soapEx);
            }

            if (Common.IsRequirementEnabled(7851, this.Site))
            {
                this.Site.CaptureRequirementIfIsFalse(
                    isSoapFaultExist,
                    7851,
                    @"[In UndoCheckOut operation] Implementation does not return a SOAP fault if the pageUrl is an invalid URL. <80> Section 3.1.4.26: wss3 does not return a SOAP fault.");
            }

            if (Common.IsRequirementEnabled(7852, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7852. SOAP fault {0}.", isSoapFaultExist ? "is returned with error string '" + errorString + "'" : "is not returned");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R7852
                bool isVerifiedR7852 = isSoapFaultExist && string.Equals(errorString, "Invalid URI: The format of the URI could not be determined.");

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR7852,
                    7852,
                    @"[In UndoCheckOut operation] Implementation does return a SOAP fault with the error string: ""Invalid URI: The format of the URI could not be determined."" if the pageUrl is an invalid URL.(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.)");
            }

            #endregion

            #region Undo the checkout with the pageUrl is an empty string, try to capture R784 and R1962
            isSoapFaultExist = false;
            errorCode = string.Empty;
            try
            {
                string nullPageUrl = string.Empty;
                this.listwsInstance.UndoCheckOut(nullPageUrl);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture R784 if the error code 0x82000001 is returned.
            Site.CaptureRequirementIfIsTrue(
                                        isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                        784,
                                        @"[In UndoCheckOut operation] If the pageUrl is an empty string the protocol server MUST return a SOAP fault with error code 0x82000001.");

            // Capture R1962 if the error code 0x82000001 is returned.
            Site.CaptureRequirementIfIsTrue(
                                        isSoapFaultExist && errorCode.Equals("0x82000001", StringComparison.OrdinalIgnoreCase),
                                        1962,
                                        @"[In UndoCheckOut operation] [If the pageUrl is an empty string the protocol server MUST return a SOAP fault with error code 0x82000001.] This indicates that the parameter pageUrl is missing or invalid.");
            #endregion

            #region Check out a file that has been checked out, if the CheckOutFile should return false, capture R2305;

            bool isUndoCheckOutSuccess = false;
            isUndoCheckOutSuccess = this.listwsInstance.UndoCheckOut(absoluteFileUrl);
            Site.Assert.IsTrue(isUndoCheckOutSuccess, "The UndoCheckOutFile operation must be successful in the first time.");

            // Check out a file that have been undo checked out
            isUndoCheckOutSuccess = this.listwsInstance.UndoCheckOut(absoluteFileUrl);

            // Capture R2306 if the UndoCheckOut returns a false value.
            Site.CaptureRequirementIfIsFalse(
                isUndoCheckOutSuccess,
                2306,
                @"[UndoCheckOutResponse] [The value is True if the operation is successful;]otherwise, False is returned.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the UndoCheckOut operation when all its input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S05_TC10_UndoCheckOut_Succeed()
        {
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            string absoluteFileUrl = this.sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            #region Check out the added file.

            bool isFicheckOutSuccessfully = false;
            isFicheckOutSuccessfully = this.listwsInstance.CheckOutFile(absoluteFileUrl, bool.TrueString, string.Empty);
            if (!isFicheckOutSuccessfully)
            {
                Site.Assert.Fail("Expect CheckOutFile operation failed due to unexpected reason.");
            }

            #endregion

            #region Undo the checkout.
            bool isUndoCheckOutSucceed = false;

            // Undo a checkOut process by calling UndoCheckOut operation.
            isUndoCheckOutSucceed = this.listwsInstance.UndoCheckOut(absoluteFileUrl);
            #endregion

            #region Capture R7781 R787 and R1973 if the UndoCheckOut succeeds and returns true value.
            // Verify requirement R7781.
            // If undo check out is successful, it means implementation does support this UndoCheckOut method. R7781 can be captured.
            if (Common.IsRequirementEnabled(7781, this.Site))
            {
                Site.CaptureRequirementIfIsTrue(
                    isUndoCheckOutSucceed,
                    7781,
                    @"Implementation does support this method[UndoCheckOut]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            Site.CaptureRequirementIfIsTrue(
                isUndoCheckOutSucceed,
                787,
                @"[In UndoCheckOut operation] If there are no errors, the protocol server MUST undo the checkout operation on the specified document.");

            Site.CaptureRequirementIfIsTrue(
                isUndoCheckOutSucceed,
                1973,
                @"[UndoCheckOutResponse]The value is True if the operation is successful;");

            #endregion
        }

        #endregion

        #endregion

        #region Override methods
        /// <summary>
        /// This method will run before test case executes
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.listwsInstance = this.Site.GetAdapter<IMS_LISTSWSAdapter>();

            Common.CheckCommonProperties(this.Site, true);

            this.sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();

            #region New initialization
            if (!TestSuiteHelper.GuardEnviromentClean())
            {
                Site.Debug.Fail("The test environment is not clean, refer the log files for details.");
            }

            // Initialize the TestSuiteHelper
            TestSuiteHelper.Initialize(this.Site, this.listwsInstance);
            #endregion
        }

        /// <summary>
        /// This method will run after test case executes
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            #region New clean up
            TestSuiteHelper.CleanUp();
            #endregion
        }

        #endregion
    }
}