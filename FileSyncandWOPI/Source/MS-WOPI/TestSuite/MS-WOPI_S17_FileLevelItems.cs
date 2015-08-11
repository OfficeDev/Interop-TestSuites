namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the operations' behaviors on file level whether follow the Open Spec definitions.
    /// </summary>
    [TestClass]
    public class MS_WOPI_S17_FileLevelItems : TestSuiteBase
    {
        #region Test Class level initialization
        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        #region Scenario 17

        /// <summary>
        /// This test case is used to verify CheckFileInfo operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC01_CheckFileInfo()
        {
            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file 
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            // Get the status code from the response of CheckFileInfo.
            int statusCode = responseOfCheckFileInfo.StatusCode;

            // Verify MS-WOPI requirement: MS-WOPI_R269
            this.Site.CaptureRequirementIfAreEqual<int>(
                          200,
                          statusCode,
                          269,
                          @"[In CheckFileInfo] Status code ""200"" means ""Success"".");

            // Verify MS-WOPI requirement: MS-WOPI_R254
            // If the CheckFileInfo instance is not null, this requirement should be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFileInfo,
                          254,
                          @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""CheckFileInfo"" is used for ""Returns information about a file.");

            // Verify MS-WOPI requirement: MS-WOPI_R263
            // If the CheckFileInfo instance is not null, this requirement should be covered.
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFileInfo,
                          263,
                          @"[In CheckFileInfo] Return information about the file and permissions that the current user has relative to that file.");

            // The URI in "CheckFileInfo" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R264
            // Verify MS-WOPI requirement: MS-WOPI_R264
            this.Site.CaptureRequirement(
                          264,
                          @"[In CheckFileInfo] HTTP Verb: GET
                          URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

            // Get the file contents
            string wopiTargetFileContentLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileContentLevelUrl);
            WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiTargetFileContentLevelUrl, commonHeaders, null);

            // Get the file the actual content.
            byte[] fileContents = WOPIResponseHelper.GetContentFromResponse(responseOfGetFile);

            if (null == fileContents || 0 == fileContents.Length)
            {
                this.Site.Assert.Fail("Could not get the file contents for the file[{0}]", fileUrl);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R325
            this.Site.CaptureRequirementIfAreEqual<int>(
                          checkFileInfo.Size,
                          fileContents.Length,
                          325,
                          @"[In Response Body] Size: The size of the file expressed in bytes.");

            if (!string.IsNullOrEmpty(checkFileInfo.UserFriendlyName))
            {
                // Verify MS-WOPI requirement: MS-WOPI_R344
                bool isVerifiedR344 = checkFileInfo.UserFriendlyName.IndexOf(Common.GetConfigurationPropertyValue("UserName", this.Site), StringComparison.OrdinalIgnoreCase) >= 0 ||
                    checkFileInfo.UserFriendlyName.IndexOf(Common.GetConfigurationPropertyValue("UserFriendlyName", this.Site), StringComparison.OrdinalIgnoreCase) >= 0;

                this.Site.CaptureRequirementIfIsTrue(
                              isVerifiedR344,
                              344,
                              @"[In Response Body] UserFriendlyName: A string that is the name of the user.");
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call CheckFileInfo operation, 
        /// the state code 404 should be returned.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC02_CheckFileInfo_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;
            try
            {
                // Return information about the file.
                WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);
                statusCode = responseOfCheckFileInfo.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R271
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          271,
                          @"[In CheckFileInfo] Status code ""404"" means
                          ""File unknown/User Unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the value of SupportsCobalt is true when call CheckFileInfo operation, the MS-WOPI server
        /// should support ExecuteCellStorageRequest and ExcecuteCellStorageRelativeRequest.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC03_CheckFileInfo_SupportsCobalt_True()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 961, this.Site))
            {
                Site.Assume.Inconclusive(@"The implementation does not support the operations ""ExecuteCellStorageRequest"" and ""ExecuteCellStorageRelativeRequest"" (does not support ""Cobalt"" feature). It is determined using SHOULDMAY PTFConfig property named R961Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            // Verify requirement MS-WOPI_R776 and MS-WOPI_R961
            // Verify MS-WOPI requirement: MS-WOPI_R961
            this.Site.CaptureRequirementIfIsTrue(
                          checkFileInfo.SupportsCobalt,
                          961,
                          @"[In WOPI Protocol Server Details]Implementation does support ExecuteCellStorageRequest (see section 3.3.5.1.7) and ExcecuteCellStorageRelativeRequest (see section 3.3.5.1.8) operations.(Microsoft SharePoint Foundation 2013 and above follow this behavior)");

            // Verify MS-WOPI requirement: MS-WOPI_R776
            this.Site.CaptureRequirementIfIsTrue(
                          checkFileInfo.SupportsCobalt,
                          776,
                          @"[In Response Body] If the value of SupportsCobalt is true, indicates that the WOPI server supports ExecuteCellStorageRequest (see section 3.3.5.1.7) and ExcecuteCellStorageRelativeRequest (see section 3.3.5.1.8) operations for this file.");
        }

        /// <summary>
        /// This test case is used to verify if the value of SupportsFolders is true when call CheckFileInfo operation, the MS-WOPI server
        /// supports EnumerateChildren and DeleteFile operations for this file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC04_CheckFileInfo_SupportsFolders_True()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            if (!checkFileInfo.SupportsFolders)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of SupportsFolders is true.");
            }

            #region call EnumerateChildren

            // Get the folder URL.
            string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

            // Get the WOPI URL.
            WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);
            string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get folder content URL.
            string wopiFolderContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFolderUrl, WOPISubResourceUrlType.FolderChildrenLevel);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFolderContentsLevelUrl);

            // Return the contents of a folder on the WOPI server.
            WopiAdapter.EnumerateChildren(wopiFolderContentsLevelUrl, commonHeaders);

            #endregion

            #region Call DeleteFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R778
            // The value of SupportsFolders is true, and EnumerateChildren and DeleteFile can be called successfully, this requirement can be covered directly.
            this.Site.CaptureRequirement(
                          778,
                          @"[In Response Body] The value of SupportsFolders is true,indicates that the WOPI server supports EnumerateChildren (see section 3.3.5.4.1) and DeleteFile (see section 3.3.5.1.9) operations for this file.");
        }

        /// <summary>
        /// This test is used to verify if the value of SupportsLocks is true when call CheckFileInfo operation, the MS-WOPI server supports
        /// Lock, Unlock, RefreshLock, and UnlockAndRelock operations for this file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC05_CheckFileInfo_SupportsLocks_True()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call CheckFileInfo

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            #endregion

            if (!checkFileInfo.SupportsLocks)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of SupportsLocks is true.");
            }

            #region Call Lock operation

            string lockIndentifier = Guid.NewGuid().ToString("N");

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Take a lock for editing a file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIndentifier);
            #endregion

            bool isExceptionThrown = false;
            string newLockIndentifier = Guid.NewGuid().ToString("N");
            try
            {
                #region Call RefreshLock operation

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                // Refresh an existing lock for modifying a file.
                WopiAdapter.RefreshLock(wopiTargetFileUrl, commonHeaders, lockIndentifier);

                #endregion

                #region Call UnlockAndRelock operation

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                // Release and retake a lock for editing a file.
                WopiAdapter.UnlockAndRelock(wopiTargetFileUrl, commonHeaders, newLockIndentifier, lockIndentifier);
                #endregion
            }
            catch (Exception)
            {
                isExceptionThrown = true;
                throw;
            }
            finally
            {
                if (isExceptionThrown)
                {
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIndentifier);
                }
            }

            #region Call UnLock

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Release a lock for editing a file.
            WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, newLockIndentifier);

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R780
            // And implement the Lock, Unlock, RefreshLock and UnlockAndRelock operations successfully.
            bool isVerifiedR780 = checkFileInfo.SupportsLocks;

            this.Site.CaptureRequirementIfIsTrue(
                          isVerifiedR780,
                          780,
                          @"[In Response Body] If the value of SupportsLocks is true, indicates that the WOPI server supports Lock (see section 3.3.5.1.3), Unlock (see section 3.3.5.1.4), RefreshLock (see section 3.3.5.1.5), and UnlockAndRelock (see section 3.3.5.1.6) operations for this file.");
        }

        /// <summary>
        /// This test case is used to verify if the value of SupportsUpdate is true when call CheckFileInfo operation, the MS-WOPI server
        /// supports PutFile and PutRelativeFile operations for this file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC06_CheckFileInfo_SupportsUpdate_True()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT(true);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call CheckFileInfo

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            #endregion

            if (!checkFileInfo.SupportsUpdate)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of SupportsUpdate is true.");
            }

            #region Call PutFile

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Update a file on the WOPI server.
            WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);

            #endregion

            #region Call PutRelativeFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
            string fileName = this.GetUniqueFileNameForPutRelatived();

            // Create a new file on the WOPI server based on the current file.
            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");

            // Create a new file on the WOPI server based on the current file.
            WopiAdapter.PutRelativeFile(
                wopiTargetFileUrl,
                commonHeaders,
                null,
                fileName,
                fileContents,
                false,
                fileContents.Length);

            // Collect the new file for PutRelativeFile.
            this.CollectNewAddedFileForPutRelativeFile(fileUrl, fileName);

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R784
            // And implement the PutFile and PutRelativeFile operations successfully.
            bool isVerifiedR784 = checkFileInfo.SupportsUpdate;

            this.Site.CaptureRequirementIfIsTrue(
                          isVerifiedR784,
                          784,
                          @"[In Response Body] If the value of SupportsUpdate is true, indicates that the WOPI server supports PutFile (see section 3.3.5.3.2) and PutRelativeFile (see section 3.3.5.1.2) operations for this file.");
        }

        /// <summary>
        /// This test case is used to verify if the value of UserCanNotWriteRelative is false when call CheckFileInfo operation, the user
        /// have sufficient permissions to create new files on the MS-WOPI server.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC07_CheckFileInfo_UserCanNotWriteRelative_false()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT(true);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            if (checkFileInfo.UserCanNotWriteRelative)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of UserCanNotWriteRelative is false.");
            }

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");

            string fileName = this.GetUniqueFileNameForPutRelatived();

            // Create a new file on the WOPI server based on the current file.
            WopiAdapter.PutRelativeFile(
                wopiTargetFileUrl,
                commonHeaders,
                null,
                fileName,
                fileContents,
                false,
                fileContents.Length);

            // Collect the new file for PutRelativeFile.
            this.CollectNewAddedFileForPutRelativeFile(fileUrl, fileName);

            // Verify MS-WOPI requirement: MS-WOPI_R922
            bool isVerifiedR922 = checkFileInfo.UserCanNotWriteRelative;

            this.Site.CaptureRequirementIfIsFalse(
                          isVerifiedR922,
                          922,
                          @"[In Response Body]If the value of UserCanNotWriteRelative is false, indicates the user  have sufficient permissions to create new files on the WOPI server.");
        }

        /// <summary>
        /// This test case is used to verify if the value of SupportsSecureStore is true when call CheckFileInfo operation,
        /// the server should support calls to a secure data store utilizing credentials stored in the file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC08_CheckFileInfo_SupportsSecureStore_true()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 963, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not support the operation ""ReadSecureStore"". It is determined using SHOULDMAY PTFConfig property named R963Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call CheckFileInfo

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            #endregion

            #region Call ReadSecureStore

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Access the WOPI server's implementation of a secure store.
            WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithGroupAndNotWindows", this.Site));

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R924
            this.Site.CaptureRequirementIfIsTrue(
                          checkFileInfo.SupportsSecureStore,
                          924,
                          @"[In Response Body] If the value of SupportsSecureStore is true, indicates that the WOPI server supports calls to a secure data store utilizing credentials stored in the file.");
        }

        /// <summary>
        /// This test case is used to verify if the value of ReadOnly is false when call CheckFileInfo operation, for this user, the file can be changed.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC09_CheckFileInfo_ReadOnly_False()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call CheckFileInfo

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            #endregion

            if (checkFileInfo.ReadOnly)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of ReadOnly is false.");
            }

            #region Call PutFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            // Update a file on the WOPI server.
            WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);

            #endregion

            #region Call GetFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Get a file with invalid token.
            WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

            string actualString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfGetFile);
            this.Site.Assert.AreEqual(exceptedUpdateContent, actualString, "Update the file should succeed.");

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R927
            bool isVerifiedR927 = checkFileInfo.ReadOnly;

            this.Site.CaptureRequirementIfIsFalse(
                          isVerifiedR927,
                          927,
                          @"[In Response Body] If the value of ReadOnly is false, indicates that, for this user, the file can be changed.");
        }

        /// <summary>
        /// This test case is used to verify if the value of UserCanWrite is true when call CheckFileInfo operation,
        /// the user has permissions to alter the file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC10_CheckFileInfo_UserCanWrite_True()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call CheckFileInfo

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            #endregion

            if (!checkFileInfo.UserCanWrite)
            {
                this.Site.Assume.Inconclusive("Test case is executed only when the value of UserCanWrite is false.");
            }

            #region Call PutFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            // Update a file on the WOPI server.
            WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);

            #endregion

            #region Call GetFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Get a file with invalid token.
            WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

            string actualString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfGetFile);
            this.Site.Assert.AreEqual(exceptedUpdateContent, actualString, "Update the file should succeed.");

            #endregion

            // Verify MS-WOPI requirement: MS-WOPI_R929
            bool isVerifiedR929 = checkFileInfo.UserCanWrite;

            this.Site.CaptureRequirementIfIsTrue(
                          isVerifiedR929,
                          929,
                          @"[In Response Body] If the value of UserCanWrite is true, indicates that the user has permissions to alter the file.");
        }

        /// <summary>
        /// This test case is used to verify the version values must not repeat for a given file.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC11_CheckFileInfo_Version()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            string oldVersion = checkFileInfo.Version;

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Update a file on the WOPI server.
            WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);

            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file again.
            responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            string newVersion = checkFileInfo.Version;

            // Verify MS-WOPI requirement: MS-WOPI_R931
            this.Site.CaptureRequirementIfAreNotEqual(
                          oldVersion,
                          newVersion,
                          931,
                          @"[In Response Body] [Version] Implementation does uniquely for a sample of N (default N=2) identify the version of the file.");
        }

        /// <summary>
        /// This test case is used to verify the value of the HostAuthenticationId from the response of
        /// CheckFileInfo operation is unique.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC12_CheckFileInfo_HostAuthenticationId()
        {
            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonCheckFileInfo = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json to object.
            CheckFileInfo checkFileBeforeSiwtchUser = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonCheckFileInfo);

            // Get the WOPI URL.
            string wopiTargetFileUrlOtherUser = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(
                fileUrl,
                WOPIRootResourceUrlType.FileLevel,
                Common.GetConfigurationPropertyValue("UserName1", this.Site),
                Common.GetConfigurationPropertyValue("Password1", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site));

            // Get the common header.
            WebHeaderCollection commonHeadersOtherUser = HeadersHelper.GetCommonHeaders(wopiTargetFileUrlOtherUser);

            // Return information about the file.
            WOPIHttpResponse responseOfCheckFileInfoOtherUser = WopiAdapter.CheckFileInfo(wopiTargetFileUrlOtherUser, commonHeadersOtherUser, null);

            // Get the json string from the response of CheckFileInfo.
            jsonCheckFileInfo = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfoOtherUser);

            // Convert the json to object.
            CheckFileInfo checkFileAfterSiwtchUser = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonCheckFileInfo);

            this.Site.CaptureRequirementIfAreNotEqual(
                          checkFileBeforeSiwtchUser.HostAuthenticationId,
                          checkFileAfterSiwtchUser.HostAuthenticationId,
                          952,
                          @"[In Response Body] HostAuthenticationId: A string that is used by the WOPI server to uniquely for a sample of N (default N=2) identify the users.");
        }

        /// <summary>
        /// This test case is used to verify PutRelativeFile operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC13_PutRelativeFile()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT(true);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");
            string fileName = this.GetUniqueFileNameForPutRelatived();

            // Create a new file on the WOPI server based on the current file.
            WOPIHttpResponse responseOfPutRelativeFile = WopiAdapter.PutRelativeFile(
                                                                wopiTargetFileUrl,
                                                                commonHeaders,
                                                                null,
                                                                fileName,
                                                                fileContents,
                                                                false,
                                                                fileContents.Length);

            // Collect the new file for PutRelativeFile.
            this.CollectNewAddedFileForPutRelativeFile(fileUrl, fileName);

            // Get the json string from the response of PutRelativeFile.
            string jsonStringForPutRelativeFile = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfPutRelativeFile);

            // Convert the json string to object.
            PutRelativeFile putRelativeFile = WOPISerializerHelper.JsonToObject<PutRelativeFile>(jsonStringForPutRelativeFile);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(putRelativeFile.Url, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Get a file.
            WOPIHttpResponse httpWebResponseForGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

            // Read message from response of GetFile.
            string actualUpdateContent = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForGetFile);

            int statusCode = httpWebResponseForGetFile.StatusCode;

            // Verify MS-WOPI requirement: MS-WOPI_R353
            this.Site.CaptureRequirementIfAreEqual(
                          "Test Put Relative file.",
                          actualUpdateContent,
                          353,
                          @"[In PutRelativeFile] Create a new file on the WOPI server based on the current file.");

            // The URI in "GetFile" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>/contents?access_token=<token>" pattern, if the operation execute successfully, capture R255 and R354
            // Verify MS-WOPI requirement: MS-WOPI_R255
            this.Site.CaptureRequirement(
                          255,
                          @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""PutRelativeFile"" is used for ""Creates a copy of a file on the WOPI server"".");

            // Verify MS-WOPI requirement: MS-WOPI_R354
            this.Site.CaptureRequirement(
                          354,
                          @"[In PutRelativeFile] HTTP Verb: POST
                          URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

            // Verify MS-WOPI requirement: MS-WOPI_R372
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCode,
                          372,
                          @"[In PutRelativeFile] Status code ""200"" means ""Success"".");

            // Verify MS-WOPI requirement: MS-WOPI_R384
            this.Site.CaptureRequirement(
                          384,
                          @"[In Response Body] URL: A URI that is the WOPI server URI of the newly created file in the form: 
                          HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

            // Call PutRelativeFile operation with an existing file name and setting header "X-WOPI-OverwriteRelativeTarget" to 'true'.
            string overwriteContentValue = string.Format("Overwrite to {0} value", Guid.NewGuid().ToString("N"));
            byte[] overwriteContent = Encoding.UTF8.GetBytes(overwriteContentValue);
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
            WopiAdapter.PutRelativeFile(
                wopiTargetFileUrl,
                commonHeaders,
                null,
                fileName,
                overwriteContent,
                true,
                overwriteContent.Length);

            // Get file content URL.
            string wopiFileContentsLevelUrlAfterOverwrite = TokenAndRequestUrlHelper.GetSubResourceUrl(putRelativeFile.Url, WOPISubResourceUrlType.FileContentsLevel);

            // Get a file.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrlAfterOverwrite);
            WOPIHttpResponse httpWebResponseForGetFileForOverwrite = WopiAdapter.GetFile(wopiFileContentsLevelUrlAfterOverwrite, commonHeaders, null);

            // Read message from response of GetFile.
            string actualContentForOverwrite = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForGetFileForOverwrite);

            // Verify the actual file content whether equal to the overwrite file content value.
            bool isEqualToOverwriteContentValue = overwriteContentValue.CompareStringValueIgnoreCase(actualContentForOverwrite, this.Site);

            // Verify MS-WOPI requirement: MS-WOPI_R974
            this.Site.CaptureRequirementIfIsTrue(
                          isEqualToOverwriteContentValue,
                          974,
                          @"[In PutRelativeFile] If the X-WOPI-OverwriteRelativeTarget is a true value and the file exists, the host MUST overwrite the file.");
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call PutRelativeFile operation, the server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC14_PutRelativeFile_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT(true);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");
            string fileName = this.GetUniqueFileNameForPutRelatived();

            int statusCode;
            try
            {
                // Create a new file on the WOPI server based on the current file.
                WOPIHttpResponse responseOfPutRelativeFile = WopiAdapter.PutRelativeFile(
                                                                    wopiTargetFileUrl,
                                                                    commonHeaders,
                                                                    null,
                                                                    fileName,
                                                                    fileContents,
                                                                    false,
                                                                    fileContents.Length);
                statusCode = responseOfPutRelativeFile.StatusCode;

                // Collect the new file for PutRelativeFile.
                this.CollectNewAddedFileForPutRelativeFile(fileUrl, fileName);
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R375
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          375,
                          @"[In PutRelativeFile] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the file is already exist when call PutRelativeFile operation, server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC15_PutRelativeFile_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the file name.
            string fileName = TestSuiteHelper.GetFileNameFromFullUrl(fileUrl);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            #region Call PutRelativeFile file name is exists.

            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");

            int statusCodeFileNameExists = 0;
            try
            {
                // Create a new file on the WOPI server based on the current file.
                WOPIHttpResponse responseOfPutRelativeFile = WopiAdapter.PutRelativeFile(
                                                                    wopiTargetFileUrl,
                                                                    commonHeaders,
                                                                    null,
                                                                    fileName,
                                                                    fileContents,
                                                                    false,
                                                                    fileContents.Length);

                statusCodeFileNameExists = responseOfPutRelativeFile.StatusCode;

                // Collect the new file for PutRelativeFile.
                this.CollectNewAddedFileForPutRelativeFile(fileUrl, fileName);
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCodeFileNameExists = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            #endregion

            // If the protocol server return status code "409" with setting header "X-WOPI-RelativeTarget" to an existing file name, that means the WOPI server does not allow to overwrite the existing file when the header "X-WOPI-OverwriteRelativeTarget" is set to false.  
            // Verify MS-WOPI requirement: MS-WOPI_R975
            this.Site.CaptureRequirementIfAreEqual(
                          409,
                          statusCodeFileNameExists,
                          975,
                          @"[In PutRelativeFile] If the X-WOPI-OverwriteRelativeTarget is a false value and the file exists, the host MUST return the status code of 409.");

            // Verify MS-WOPI requirement: MS-WOPI_R376
            this.Site.CaptureRequirementIfAreEqual(
                          409,
                          statusCodeFileNameExists,
                          376,
                          @"[In PutRelativeFile] Status code ""409"" means ""Target file already exists"".");
        }

        /// <summary>
        /// This test case is used to verify if X-WOPI-SuggestedTarget extension is provided,
        /// the name of the initial file without extension is combined with the extension to
        /// create the proposed name.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC16_PutRelativeFile_SuggestedTarget()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 788, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not combine the name of the initial file without extension with the extension specified in X-WOPI-SuggestedTarget to create the proposed name, if only the extension is provided in X-WOPI-SuggestedTarget in PutRelativeFile operation. It is determined using SHOULDMAY PTFConfig property named R788Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", this.Site);

            // Get the file name.
            string fileName = TestSuiteHelper.GetFileNameFromFullUrl(fileUrl);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            byte[] fileContents = Encoding.UTF8.GetBytes("Test Put Relative file.");

            // Create a new file on the WOPI server based on the current file.
            WOPIHttpResponse responseOfPutRelativeFile = WopiAdapter.PutRelativeFile(
                                                                wopiTargetFileUrl,
                                                                commonHeaders,
                                                                ".wopiTest",
                                                                null,
                                                                fileContents,
                                                                false,
                                                                fileContents.Length);

            // Collect the new file for PutRelativeFile.
            string expectedFileName = fileName.Substring(0, fileName.IndexOf(".")) + ".wopiTest";
 
            // Read message from response of GetFile.
            string actualUpdateContent = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfPutRelativeFile);

            // Convert the json to object.
            PutRelativeFile putRelativeFile = WOPISerializerHelper.JsonToObject<PutRelativeFile>(actualUpdateContent);

            // Verify requirement MS-WOPI_R788
            bool isEqualToExpectFileName = expectedFileName.CompareStringValueIgnoreCase(putRelativeFile.Name, this.Site);

            // Collect the added file according to the actual added file name.
            if (isEqualToExpectFileName)
            {
                this.CollectNewAddedFileForPutRelativeFile(fileUrl, expectedFileName);
            }
            else
            {
                if (!string.IsNullOrEmpty(putRelativeFile.Name))
                {
                    this.CollectNewAddedFileForPutRelativeFile(fileUrl, putRelativeFile.Name);
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                          isEqualToExpectFileName,
                          788,
                          @"[In PutRelativeFile] [X-WOPI-SuggestedTarget] Implementation does support the name of the initial file without extension is combined with the extension to create the proposed name,if only the extension is provided.(Microsoft SharePoint Foundation 2013 and above follow this behavior)");
        }

        /// <summary>
        /// This test case is used to verify Lock operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC17_Lock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCodeOfLock;
            string lockIdentifierValue = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WOPIHttpResponse httpWebResponseForLock = WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);

            bool isWebExceptionRaise = false;
            bool isDeleteFileSuccessful = false;
            try
            {
                statusCodeOfLock = httpWebResponseForLock.StatusCode;

                // Verify MS-WOPI requirement: MS-WOPI_R405
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfLock,
                              405,
                              @"[In Lock] Status code ""200"" means ""Success"".");

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                try
                {
                    // Delete this file which is locked.
                    WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

                    // The file has been deleted, so remove it from the clean up list.
                    this.ExcludeFileFromTheCleanUpProcess(fileUrl);
                    isDeleteFileSuccessful = true;
                }
                catch (WebException webEx)
                {
                    isWebExceptionRaise = true;
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R256
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              256,
                              @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""Lock"" is used for ""Takes a lock for editing a file"".");

                // Verify MS-WOPI requirement: MS-WOPI_R390
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              390,
                              @"[In Lock] Take a lock for editing a file.");

                // The URI in "Lock" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R660
                // Verify MS-WOPI requirement: MS-WOPI_R391
                this.Site.CaptureRequirement(
                              391,
                              @"[In Lock] HTTP Verb: POST
                              URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

                // If the WOPI server perform the LOCK operation successfully with the specified lock indentifier in X-WOPI-Lock header, then capture R397, R411
                // Verify MS-WOPI requirement: MS-WOPI_R397
                this.Site.CaptureRequirement(
                              397,
                              @"[In Lock] X-WOPI-Lock is a string provided by the WOPI client that the WOPI server MUST use to identify the lock on the file.");

                // Verify MS-WOPI requirement: MS-WOPI_R411
                this.Site.CaptureRequirement(
                              411,
                              @"[In Processing Details] The WOPI server MUST use the string provided in the X-WOPI-Lock header to create a lock on a file.");
            }
            finally
            {
                if (!isDeleteFileSuccessful)
                {
                    // Release a lock for editing a file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call Lock operation, the server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC18_Lock_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // If the file has been deleted successfully by calling DeleteFile operation, remove it from the clean up process.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            bool lockSuccessfully = false;
            string lockIdentifierValue = Guid.NewGuid().ToString("N");

            try
            {
                try
                {
                    // Take a lock for editing a file.
                    WOPIHttpResponse httpWebResponseForLock = WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);
                    statusCode = httpWebResponseForLock.StatusCode;
                    lockSuccessfully = true;
                }
                catch (WebException webEx)
                {
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R407
                this.Site.CaptureRequirementIfAreEqual(
                              404,
                              statusCode,
                              407,
                              @"[In Lock] Status code ""404"" means ""File unknown/User unauthorized"".");
            }
            finally
            {
                if (lockSuccessfully)
                {
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is locked when call Lock operation, the server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC19_Lock_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string lockIdentifierValue = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WOPIHttpResponse httpWebResponseForLock = WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);

            string newLockIdentifierValue = Guid.NewGuid().ToString("N");
            bool isLockAgainSuccessful = false;
            try
            {
                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                int statusCode = 0;
                try
                {
                    // Take a lock for editing a file.
                    httpWebResponseForLock = WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, newLockIdentifierValue);
                    statusCode = httpWebResponseForLock.StatusCode;
                    isLockAgainSuccessful = true;
                }
                catch (WebException webEx)
                {
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R408
                this.Site.CaptureRequirementIfAreEqual(
                              409,
                              statusCode,
                              408,
                              @"[In Lock] Status code ""409"" means ""Lock mismatch"".");
            }
            finally
            {
                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                string currentLockIndentifier = isLockAgainSuccessful ? newLockIdentifierValue : lockIdentifierValue;
                WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, currentLockIndentifier);
            }
        }

        /// <summary>
        /// This test case is used to verify Unlock operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC20_UnLock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Take a lock for editing a file.
            string lockIdentifierValue = Guid.NewGuid().ToString("N");
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);

            bool isDeleteFileFirstTimeSuccessful = false;
            bool isDeleteFileSecondTimeSuccessful = false;
            try
            {
                int statusCodeOfDelete = 0;
                try
                {
                    // Delete this file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

                    // The file has been deleted, so remove it from the clean up list.
                    isDeleteFileFirstTimeSuccessful = true;
                    this.ExcludeFileFromTheCleanUpProcess(fileUrl);
                }
                catch (WebException webEx)
                {
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCodeOfDelete = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                this.Site.Assert.AreEqual(404, statusCodeOfDelete, "The file is locked when call DeleteFile operation, so the error should be 404.");

                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WOPIHttpResponse httpWebResponseForUnLock = WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);
                int statusCodeOfUnLocked = httpWebResponseForUnLock.StatusCode;

                // Verify MS-WOPI requirement: MS-WOPI_R424
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfUnLocked,
                              424,
                              @"[In Unlock] Status code ""200"" means ""Success"".");

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                // Delete this file.
                WOPIHttpResponse httpWebResponseAfterUnlockDeleteFile = WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);
                statusCodeOfDelete = httpWebResponseAfterUnlockDeleteFile.StatusCode;

                // The file has been deleted, so remove it from the clean up list.
                this.ExcludeFileFromTheCleanUpProcess(fileUrl);
                isDeleteFileSecondTimeSuccessful = true;

                // If the WOPI server perform the UNLOCK operation successfully with specified lock indentifier in header "X-WOPI-Lock", then capture R420, R257, R414, R415
                // Verify MS-WOPI requirement: MS-WOPI_R420
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfDelete,
                              420,
                              @"[In Unlock] X-WOPI-Lock is a string provided by the WOPI client that the WOPI server MUST use to identify the lock on the file.");

                // Verify MS-WOPI requirement: MS-WOPI_R257
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfDelete,
                              257,
                              @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""Unlock"" is used for ""Releases a lock for editing a file"".");

                // Verify MS-WOPI requirement: MS-WOPI_R414
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfDelete,
                              414,
                              @"[In Unlock] Release a lock for editing a file.");

                // The URI in "Unlock" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R415
                // Verify MS-WOPI requirement: MS-WOPI_R415
                this.Site.CaptureRequirement(
                              415,
                              @"[In Unlock] HTTP Verb: POST
                              URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");
            }
            finally
            {
                if (!isDeleteFileFirstTimeSuccessful && !isDeleteFileSecondTimeSuccessful)
                {
                    // Release a lock for editing a file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call UnLock operation,
        /// the server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC21_UnLock_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;
            try
            {
                // Release a lock for editing a file.
                WOPIHttpResponse httpWebResponseForUnLock = WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"));
                statusCode = httpWebResponseForUnLock.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R426
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          426,
                          @"[In Unlock] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the file is locked When call Unlock operation with new GUID as a lock identifier, the server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC22_UnLock_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string identifierForLock = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, identifierForLock);

            int statusCode = 0;
            bool isUnlockSuccessful = false;
            try
            {
                try
                {
                    // Take a unlock for editing a file with new GUID as a lock identifier
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WOPIHttpResponse httpWebResponseForUnLock = WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"));
                    statusCode = httpWebResponseForUnLock.StatusCode;
                    isUnlockSuccessful = true;
                }
                catch (WebException webEx)
                {
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R427
                this.Site.CaptureRequirementIfAreEqual(
                              409,
                              statusCode,
                              427,
                              @"[In Unlock] Status code ""409"" means ""Lock mismatch"".");
            }
            finally
            {
                if (!isUnlockSuccessful)
                {
                    // Release a lock for editing a file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, identifierForLock);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify the RefreshLock operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC23_RefreshLock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
            string lockIndentifier = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIndentifier);

            bool isDeleteFileSuccessful = false;
            try
            {
                // Refresh an existing lock for modifying a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WOPIHttpResponse httpWebResponseForRefreshLock = WopiAdapter.RefreshLock(wopiTargetFileUrl, commonHeaders, lockIndentifier);
                int statusCodeOfRefreshLock = httpWebResponseForRefreshLock.StatusCode;

                // Verify MS-WOPI requirement: MS-WOPI_R441
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfRefreshLock,
                              441,
                              @"[In RefreshLock] Status code ""200"" means ""Success"".");

                bool isWebExceptionRaise = false;
                try
                {
                    // Delete this file which is refresh locked.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

                    // The file has been deleted, so remove it from the clean up list.
                    this.ExcludeFileFromTheCleanUpProcess(fileUrl);
                    isDeleteFileSuccessful = true;
                }
                catch (WebException webEx)
                {
                    isWebExceptionRaise = true;
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R258
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              258,
                              @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""RefreshLock"" is used for ""Refreshes a lock for editing a file"".");

                // Verify MS-WOPI requirement: MS-WOPI_R431
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              431,
                              @"[In RefreshLock] Refresh an existing lock for modifying a file.");

                // The URI in "RefreshLock" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R432
                // Verify MS-WOPI requirement: MS-WOPI_R432
                this.Site.CaptureRequirement(
                              432,
                              @"[In RefreshLock] HTTP Verb: POST
                              URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

                // If the WOPI server perform "RefreshLock" operation successfully with specified lock indentifier in heaer "X-WOPI-Lock", then capture R437
                // Verify MS-WOPI requirement: MS-WOPI_R437
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfRefreshLock,
                              437,
                              @"[In RefreshLock] X-WOPI-Lock is a string provided by the WOPI client that the WOPI server MUST use to identify the lock on the file.");
            }
            finally
            {
                if (!isDeleteFileSuccessful)
                {
                    // Release a lock for editing a file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, lockIndentifier);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call RefreshLock operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC24_RefreshLock_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            int statusCode = 0;
            try
            {
                // Refresh an existing lock for modifying a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WOPIHttpResponse httpWebResponseForRefreshLock = WopiAdapter.RefreshLock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"));
                statusCode = httpWebResponseForRefreshLock.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R443
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          443,
                          @"[In RefreshLock] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the Lock is mismatch when call RefreshLock operation,
        /// server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC25_RefreshLock_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string identifierForLock = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, identifierForLock);

            try
            {
                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                int statusCode = 0;
                try
                {
                    // Refresh a lock for the file with new GUID as lock identifier
                    WOPIHttpResponse httpWebResponseForRefreshLock = WopiAdapter.RefreshLock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"));
                    statusCode = httpWebResponseForRefreshLock.StatusCode;
                }
                catch (WebException webEx)
                {
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R444
                this.Site.CaptureRequirementIfAreEqual(
                              409,
                              statusCode,
                              444,
                              @"[In RefreshLock] Status code ""409"" means ""Lock mismatch"".");
            }
            finally
            {
                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, identifierForLock);
            }
        }

        /// <summary>
        /// This test case is used to verify UnlockAndRelock operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC26_UnlockAndRelock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string lockIdentifierValue = Guid.NewGuid().ToString("N");

            // Take a lock for editing a file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, lockIdentifierValue);

            string unlockAndRelockIdentifierValue = Guid.NewGuid().ToString("N");
            bool isDeleteFileSuccessful = false;
            bool isRelockSuccessful = false;
            try
            {
                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                int statusCodeOfUnlockAndRelock = 0;

                // Release and retake a lock for editing a file.
                WOPIHttpResponse httpWebResponseForUnLockAndRelock = WopiAdapter.UnlockAndRelock(wopiTargetFileUrl, commonHeaders, unlockAndRelockIdentifierValue, lockIdentifierValue);
                statusCodeOfUnlockAndRelock = httpWebResponseForUnLockAndRelock.StatusCode;
                isRelockSuccessful = true;

                // Verify MS-WOPI requirement: MS-WOPI_R462
                this.Site.CaptureRequirementIfAreEqual(
                              200,
                              statusCodeOfUnlockAndRelock,
                              462,
                              @"[In UnlockAndRelock] Status code ""200"" means ""Success"".");

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                bool isWebExceptionRaise = false;
                try
                {
                    // Delete this file which is refresh locked.
                    WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

                    // The file has been deleted, so remove it from the clean up list.
                    this.ExcludeFileFromTheCleanUpProcess(fileUrl);
                    isDeleteFileSuccessful = true;
                }
                catch (WebException webEx)
                {
                    isWebExceptionRaise = true;
                    HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                    this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R259
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              259,
                              @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""UnlockAndRelock"" is used for ""Releases and then retakes a lock for editing a file"".");

                // Verify MS-WOPI requirement: MS-WOPI_R449
                this.Site.CaptureRequirementIfIsTrue(
                              isWebExceptionRaise,
                              449,
                              @"[In UnlockAndRelock] Release and retake a lock for editing a file.");

                // The URI in "UnlockAndRelock" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R432
                // Verify MS-WOPI requirement: MS-WOPI_R432
                this.Site.CaptureRequirement(
                              450,
                              @"[In UnlockAndRelock] HTTP Verb: POST
                              URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

                // If the "UnlockAndRelock" execute successfully with specified new lock indentifier in header "X-WOPI-Lock" and the file could not be deleted, then capture R455 and R458
                // Verify MS-WOPI requirement: MS-WOPI_R455
                this.Site.CaptureRequirement(
                              455,
                              @"[In UnlockAndRelock] X-WOPI-Lock is a string provided by the WOPI client that the WOPI server MUST use to identify the lock on the file.");

                // Verify MS-WOPI requirement: MS-WOPI_R458
                this.Site.CaptureRequirement(
                              458,
                              @"[In UnlockAndRelock] X-WOPI-OldLock is a string previously provided by the WOPI client that the WOPI server MUST have used to identify the lock on the file.");
            }
            finally
            {
                if (!isDeleteFileSuccessful)
                {
                    // Release a lock for editing a file.
                    commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                    string currentLockIndentifier = isRelockSuccessful ? unlockAndRelockIdentifierValue : lockIdentifierValue;
                    WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, currentLockIndentifier);
                }
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call UnlockAndRelock, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC27_UnlockAndRelock_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string lockIdentifierValue = Guid.NewGuid().ToString("N");

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;
            try
            {
                // Release and retake a lock for editing a file with invalid token.
                WOPIHttpResponse httpWebResponseForUnLockAndRelock = WopiAdapter.UnlockAndRelock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"), lockIdentifierValue);
                statusCode = httpWebResponseForUnLockAndRelock.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R464
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          464,
                          @"[In UnlockAndRelock] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the file is lock mismatch when call UnlockAndRelock operation,
        /// server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC28_UnlockAndRelock_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;
            try
            {
                // Release and retake a lock for editing a file with invalid token.
                WOPIHttpResponse httpWebResponseForUnLockAndRelock = WopiAdapter.UnlockAndRelock(wopiTargetFileUrl, commonHeaders, Guid.NewGuid().ToString("N"), Guid.NewGuid().ToString("N"));
                statusCode = httpWebResponseForUnLockAndRelock.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R465
            this.Site.CaptureRequirementIfAreEqual(
                          409,
                          statusCode,
                          465,
                          @"[In UnlockAndRelock] Status code ""409"" means ""Lock mismatch"".");
        }

        /// <summary>
        /// This test case is used to verify DeleteFile operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC29_DeleteFile()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCodeOfDeleteFile = 0;

            // Delete the file.
            WOPIHttpResponse httpWebResponseForDeleteFile = WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);
            statusCodeOfDeleteFile = httpWebResponseForDeleteFile.StatusCode;

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCodeOfCheckFileInfo = 0;
            try
            {
                // Return information about the file with an invalid user.
                WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);
                statusCodeOfCheckFileInfo = responseOfCheckFileInfo.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCodeOfCheckFileInfo = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R262
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfCheckFileInfo,
                          262,
                          @"[In HTTP://server/<...>/wopi*/files/<id>] Operation ""DeleteFile"" is used for ""Removes a file from the WOPI server"".");

            // Verify MS-WOPI requirement: MS-WOPI_R510
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfCheckFileInfo,
                          510,
                          @"[In DeleteFile] Delete a file.");

            // The URI in "DeleteFile" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R511
            // Verify MS-WOPI requirement: MS-WOPI_R511
            this.Site.CaptureRequirement(
                          511,
                          @"[In DeleteFile] HTTP Verb: POST
                          URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

            // Verify MS-WOPI requirement: MS-WOPI_R517
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCodeOfDeleteFile,
                          517,
                          @"[In DeleteFile] Status code ""200"" means ""Success"".");

            // Verify MS-WOPI requirement: MS-WOPI_R523
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfCheckFileInfo,
                          523,
                          @"[In Processing Details] The WOPI server MUST delete the file if possible, given the permissions and state of the file.");
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call DeleteFile operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC30_DeleteFile_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WOPIHttpResponse httpWebResponseForDeleteFile = WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;
            try
            {
                // Delete the file again.
                httpWebResponseForDeleteFile = WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);
                statusCode = httpWebResponseForDeleteFile.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R519
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCode,
                          519,
                          @"[In DeleteFile] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify the ReadSecureStore operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC31_ReadSecureStore()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 963, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not support the operation ""ReadSecureStore"". It is determined using SHOULDMAY PTFConfig property named R963Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UploadedFileUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCode = 0;

            // Access the WOPI server's implementation of a secure store.
            WOPIHttpResponse responseOfReadSecureStore = WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithIndividualAndWindows", this.Site));
            statusCode = responseOfReadSecureStore.StatusCode;

            // Verify MS-WOPI requirement: MS-WOPI_R536
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCode,
                          536,
                          @"[In ReadSecureStore] Status code ""200"" means ""Success"".");

            // The URI in "ReadSecureStore" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R527
            // Verify MS-WOPI requirement: MS-WOPI_R527
            this.Site.CaptureRequirement(
                          527,
                          @"[In ReadSecureStore] HTTP Verb: POST
                          URI: HTTP://server/<...>/wopi*/files/<id>?access_token=<token>");

            // Verify MS-WOPI requirement: MS-WOPI_R963
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCode,
                          963,
                          @"[In WOPI Protocol Server Details]Implementation does support ReadSecureStore (see section 3.3.5.1.10) operation.(Microsoft SharePoint Server 2013 follow this behavior)");
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call ReadSecureStore operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC32_ReadSecureStore_Fail404()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 963, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not support the operation ""ReadSecureStore"". It is determined using SHOULDMAY PTFConfig property named R963Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            int statusCodeOfReadSecureStore = 0;
            try
            {
                // Access the WOPI server's implementation of a secure store.
                WOPIHttpResponse responseOfReadSecureStore = WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithIndividualAndWindows", this.Site));
                statusCodeOfReadSecureStore = responseOfReadSecureStore.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCodeOfReadSecureStore = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R538
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfReadSecureStore,
                          538,
                          @"[In ReadSecureStore] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify the value of IsWindowsCredentials, specifies that
        /// UserName corresponds to WindowsUserName and Password corresponds to WindowsPassword
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC33_ReadSecureStore_IsWindowsCredentials()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 963, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not support the operation ""ReadSecureStore"". It is determined using SHOULDMAY PTFConfig property named R963Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the target app information from the secure store, whose CredentialType is Windows credentials.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
            WOPIHttpResponse responseOfReadSecureStoreForWindowsCredential = WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithIndividualAndWindows", this.Site));

            // Get the app information from the response of ReadSecureStore.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfReadSecureStoreForWindowsCredential);
            ReadSecureStore readSStoreOfWindowsCredential = WOPISerializerHelper.JsonToObject<ReadSecureStore>(jsonString);

            // Verify MS-WOPI requirement: MS-WOPI_R544
            string valueOfUserCredentialItem = Common.GetConfigurationPropertyValue("ValueOfUserCredentialItem", this.Site);
            string valueOfPasswordCredentialItem = Common.GetConfigurationPropertyValue("ValueOfPasswordCredentialItem", this.Site);

            this.Site.Assert.AreEqual<string>(valueOfUserCredentialItem, readSStoreOfWindowsCredential.UserName, "The user credential item should match the expected value.");
            this.Site.Assert.AreEqual<string>(valueOfPasswordCredentialItem, readSStoreOfWindowsCredential.Password, "The password credential item should match the expected value.");

            // For a target app whose CredentialType is Windows credentials, the "IsWindowsCredentials" property should be 'true'
            this.Site.CaptureRequirementIfIsTrue(
                          readSStoreOfWindowsCredential.IsWindowsCredentials,
                          544,
                          @"[In Response Body] IsWindowsCredentials: A Boolean value that specifies that UserName corresponds to WindowsUserName and Password corresponds to WindowsPassword (see [MS-SSWPS] section 2.2.5.4).");
        }

        /// <summary>
        /// This test case is used to verify the value of IsGroup, specifies that the secure store application is a Group.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S17_TC34_ReadSecureStore_IsGroup()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 963, this.Site))
            {
                this.Site.Assume.Inconclusive(@"The implementation does not support the operation ""ReadSecureStore"". It is determined using SHOULDMAY PTFConfig property named R963Enabled_MS-WOPI.");
            }

            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Access the WOPI server's implementation of a secure store.
            WOPIHttpResponse responseOfReadSecureStore = WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithIndividualAndWindows", this.Site));

            // Get the json string from the response of ReadSecureStore.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfReadSecureStore);

            // Convert the json string to object.
            ReadSecureStore readSStore = WOPISerializerHelper.JsonToObject<ReadSecureStore>(jsonString);

            bool isGroupBrfore = readSStore.IsGroup;

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Call ReadSecureStore operation again.
            responseOfReadSecureStore = WopiAdapter.ReadSecureStore(wopiTargetFileUrl, commonHeaders, Common.GetConfigurationPropertyValue("IdOfAppWithGroupAndNotWindows", this.Site));

            // Get the json string from the response of ReadSecureStore.
            jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfReadSecureStore);

            // Convert the json string to object.
            readSStore = WOPISerializerHelper.JsonToObject<ReadSecureStore>(jsonString);

            bool isGroupAfter = readSStore.IsGroup;

            // Verify MS-WOPI requirement: MS-WOPI_R545
            bool isVerifiedR545 = isGroupBrfore == false && isGroupAfter == true;

            this.Site.CaptureRequirementIfIsTrue(
                          isVerifiedR545,
                          545,
                          @"[In Response Body] IsGroup: A Boolean value that specifies that the secure store application is a Group (see [MS-SSWPS] section 2.2.5.5).");
        }

        #endregion 
    }
}