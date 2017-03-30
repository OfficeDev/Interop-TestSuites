namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the operations' behaviors on file content level whether follow the Open Spec definitions.
    /// </summary>
    [TestClass]
    public class MS_WOPI_S19_FileContentLevelItems : TestSuiteBase
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

        #region Scenario 19

        /// <summary>
        /// This test case is used to verify the GetFile operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC01_GetFile()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            int statusCode = 0;

            // Get a file.
            WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);
            statusCode = responseOfGetFile.StatusCode;

            // Verify MS-WOPI requirement: MS-WOPI_R667
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCode,
                          667,
                          @"[In GetFile] Status code ""200"" means ""Success"".");

            // Verify MS-WOPI requirement: MS-WOPI_R657
            this.Site.CaptureRequirementIfIsNotNull(
                          responseOfGetFile,
                          657,
                          @"[In HTTP://server/<...>/wopi*/files/<id>/contents] Operation ""GetFile"" is used for ""Request message to retrieve a file for the HTTP://server/<...>/wopi*/files/<id>/contents operation"".");

            // Verify MS-WOPI requirement: MS-WOPI_R659
            this.Site.CaptureRequirementIfIsNotNull(
                          responseOfGetFile,
                          659,
                          @"[In GetFile] Get a file.");

            // The URI in "GetFile" WOPI request follow the "HTTP://server/<...>/wopi*/files/<id>/contents?access_token=<token>" pattern, if the operation execute successfully, capture R660
            // Verify MS-WOPI requirement: MS-WOPI_R660
            this.Site.CaptureRequirement(
                          660,
                          @"[In GetFile] HTTP Verb: GET
                          URI: HTTP://server/<...>/wopi*/files/<id>/contents?access_token=<token>");

            // Verify MS-WOPI requirement: MS-WOPI_R672
            this.Site.CaptureRequirementIfIsNotNull(
                          responseOfGetFile,
                          672,
                          @"[In Response Body] The binary contents of the file.");
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call GetFile operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC02_GetFile_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete the file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            int statusCodeOfGetFile = 0;
            try
            {
                // Get a file.
                WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);
                statusCodeOfGetFile = responseOfGetFile.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCodeOfGetFile = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R669
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfGetFile,
                          669,
                          @"[In GetFile] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify the PutFile operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC03_PutFile()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            int statusCode = 0;

            // Update a file on the WOPI server.
            WOPIHttpResponse httpWebResponseOfPutFile = WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);
            statusCode = httpWebResponseOfPutFile.StatusCode;

            // Verify MS-WOPI requirement: MS-WOPI_R687
            this.Site.CaptureRequirementIfAreEqual(
                          200,
                          statusCode,
                          687,
                          @"[In PutFile] Status code ""200"" means ""Success"".");

            if (Common.IsRequirementEnabled("MS-WOPI", 685004001, this.Site))
            {
                // Verify MS-WOPI requirement: MS-WOPI_R685004001
                this.Site.CaptureRequirementIfIsTrue(
                              string.IsNullOrEmpty(httpWebResponseOfPutFile.Headers.Get("X-WOPI-Lock")),
                              685004001,
                              @"[In PutFile] Implementation does not include the header X-WOPI-Lock when responding with the 200 status code. (SharePoint Foundation 2010 and above follows this behavior).");
            }

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            // Get a file.
            WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

            // Read message from response of GetFile.
            string actualUpdateContent = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfGetFile);

            // Verify MS-WOPI requirement: MS-WOPI_R674
            this.Site.CaptureRequirementIfAreEqual(
                          exceptedUpdateContent,
                          actualUpdateContent,
                          674,
                          @"[In PutFile] Update a file on the WOPI server.");

            // Verify MS-WOPI requirement: MS-WOPI_R675
            this.Site.CaptureRequirement(
                          675,
                          @"[In PutFile] HTTP Verb: POST
                          URI: HTTP://server/<...>/wopi*/files/<id>/contents?access_token=<token>");

            // Verify MS-WOPI requirement: MS-WOPI_R696
            this.Site.CaptureRequirementIfAreEqual(
                          exceptedUpdateContent,
                          actualUpdateContent,
                          696,
                          @"[In Processing Details] The WOPI server MUST update the binary of the file identified by <id> to match the binary contents in the request body, if the response indicates Success.");

            // Verify MS-WOPI requirement: MS-WOPI_R658
            this.Site.CaptureRequirementIfAreEqual(
                          exceptedUpdateContent,
                          actualUpdateContent,
                          658,
                          @"[In HTTP://server/<...>/wopi*/files/<id>/contents] Operation ""PutFile"" is used for 
                          ""Request message to update a file for the HTTP://server/<...>/wopi*/files/<id>/contents operation"".");

            // Verify MS-WOPI requirement: MS-WOPI_R673
            this.Site.CaptureRequirementIfAreEqual(
                          exceptedUpdateContent,
                          actualUpdateContent,
                          673,
                          @"[In Processing Details] The WOPI server MUST return the complete binary of the file identified by <id> in the response body, if the response indicates Success.");
        }

        /// <summary>
        /// This test case is used to verify if the file is unknown when call PutFile operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC04_PutFile_Fail404()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Delete this file.
            WopiAdapter.DeleteFile(wopiTargetFileUrl, commonHeaders);

            // The file has been deleted, so remove it from the clean up list.
            this.ExcludeFileFromTheCleanUpProcess(fileUrl);

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            int statusCodeOfPutFile = 0;
            try
            {
                // Update a file on the WOPI server.
                WOPIHttpResponse httpWebResponseOfPutFile = WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);
                statusCodeOfPutFile = httpWebResponseOfPutFile.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCodeOfPutFile = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R689
            this.Site.CaptureRequirementIfAreEqual(
                          404,
                          statusCodeOfPutFile,
                          689,
                          @"[In PutFile] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        /// <summary>
        /// This test case is used to verify if the file is locked when call PutFile operation,
        /// server will return 409 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC05_PutFile_Fail409()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
            string identifierForLock = Guid.NewGuid().ToString("N");

            // Lock this file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, identifierForLock);

            try
            {
                // Get file content URL.
                string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

                string exceptedUpdateContent = "WOPI PUT file test";
                byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
                string identifier = Guid.NewGuid().ToString("N");

                int statusCode = 0;
                HttpWebResponse errorResponse = null;
                try
                {
                    // Update a file on the WOPI server.
                    WOPIHttpResponse httpWebResponseOfPutFile = WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);
                    statusCode = httpWebResponseOfPutFile.StatusCode;
                }
                catch (WebException webEx)
                {
                    errorResponse = this.GetErrorResponseFromWebException(webEx);
                    statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
                }

                // Verify MS-WOPI requirement: MS-WOPI_R690
                this.Site.CaptureRequirementIfAreEqual(
                              409,
                              statusCode,
                              690,
                              @"[In PutFile] Status code ""409"" means ""Lock mismatch"".");

                if (Common.IsRequirementEnabled("MS-WOPI", 685003, this.Site))
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R685003
                    this.Site.CaptureRequirementIfIsTrue(
                                  errorResponse.Headers.Get("X-WOPI-Lock") != null,
                                  685003,
                                  @"[In PutFile] This header [X-WOPI-Lock] MUST be included when responding with the 409 status code. ");

                    // Verify MS-WOPI requirement: MS-WOPI_R980001
                    this.Site.CaptureRequirementIfIsTrue(
                                  errorResponse.Headers.Get("X-WOPI-Lock") != null,
                                  980001,
                                  @"[In PutFile] an X-WOPI-Lock response header containing the value of the current lock on the file MUST be included when using this response code [409].");

                    Boolean VerifyR685005 = false;
                    for (int i = 0; i < errorResponse.Headers.Count; i++)
                    {
                        if (errorResponse.Headers.AllKeys[i] == "X-WOPI-Lock")
                        {
                            VerifyR685005 = true;
                            break;
                        }
                    }
                    // Verify MS-WOPI requirement: MS-WOPI_R685005
                    this.Site.CaptureRequirementIfIsTrue(
                        VerifyR685005,
                        685005,
                        @"[In PutFile] X-WOPI-Lock is a string.");
                }

                if (Common.IsRequirementEnabled("MS-WOPI", 685009001, this.Site))
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R685009001
                    this.Site.CaptureRequirementIfIsTrue(
                                  string.IsNullOrEmpty(errorResponse.Headers.Get("X-WOPI-LockFailureReason")),
                                  685009001,
                                  @"[In PutFile] Implementation does not include the header X-WOPI-LockFailureReason when responding with the 409 status code. (SharePoint Foundation 2010 and above follows this behavior).");
                }

                if (Common.IsRequirementEnabled("MS-WOPI", 685012001, this.Site))
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R685012001
                    this.Site.CaptureRequirementIfIsTrue(
                                  string.IsNullOrEmpty(errorResponse.Headers.Get("X-WOPI-LockedByOtherInterface")),
                                  685012001,
                                  @"[In PutFile] Implementation does not include the header X-WOPI-LockedByOtherInterface  when responding with the 409 status code. (SharePoint Foundation 2010 and above follows this behavior).");
                }
            }
            finally
            {
                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, identifierForLock);
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is locked before call PutFile operation,
        /// MS-WOPI server MUST provide the matching lock value in order for this operation to succeed.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC06_PutFile_Lock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string identifierForLock = Guid.NewGuid().ToString("N");

            // Lock this file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, identifierForLock);

            try
            {
                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

                string exceptedUpdateContent = "WOPI PUT file test";
                byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);

                // Update a file on the WOPI server with the identifier of Lock.
                WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifierForLock);

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

                // Get a file.
                WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

                // Read message from response of GetFile.
                string actualUpdateContent = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfGetFile);

                // Verify MS-WOPI requirement: MS-WOPI_R697
                this.Site.CaptureRequirementIfAreEqual(
                              exceptedUpdateContent,
                              actualUpdateContent,
                              697,
                              @"[In Processing Details] If the file is currently associated with a lock established by the Lock operation (see section 3.3.5.1.3) [or the UnlockAndRelock operation (see section 3.3.5.1.6)] the WOPI server MUST provide the matching lock value in order for this operation to succeed.");
            }
            finally
            {
                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, identifierForLock);
            }
        }

        /// <summary>
        /// This test case is used to verify if the file is unlocked and relocked before call PutFile operation,
        /// MS-WOPI server MUST provide the matching lock value in order for this operation to succeed.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S19_TC07_PutFile_UnlockAndRelock()
        {
            // Get the file URL.
            string fileUrl = this.AddFileToSUT();

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            string identifierForLock = Guid.NewGuid().ToString("N");

            // Lock this file.
            WopiAdapter.Lock(wopiTargetFileUrl, commonHeaders, identifierForLock);

            string identifierForUnlockAndRelock = Guid.NewGuid().ToString("N");
            bool isUnlockAndRelockSuccessful = false;
            try
            {
                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

                // Release and retake a lock for editing a file.
                WopiAdapter.UnlockAndRelock(wopiTargetFileUrl, commonHeaders, identifierForUnlockAndRelock, identifierForLock);
                isUnlockAndRelockSuccessful = true;

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

                string exceptedUpdateContent = "WOPI PUT file test";
                byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);

                // Update a file on the WOPI server with the identifier of Lock.
                WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifierForUnlockAndRelock);

                // Get the common header.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

                // Get a file.
                WOPIHttpResponse responseOfGetFile = WopiAdapter.GetFile(wopiFileContentsLevelUrl, commonHeaders, null);

                // Read message from response of GetFile.
                string actualUpdateContent = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfGetFile);

                // Verify MS-WOPI requirement: MS-WOPI_R698
                this.Site.CaptureRequirementIfAreEqual(
                              exceptedUpdateContent,
                              actualUpdateContent,
                              698,
                              @"[In Processing Details] If the file is currently associated with a lock established by [the Lock operation (see section 3.3.5.1.3) or] the UnlockAndRelock operation (see section 3.3.5.1.6) the WOPI server MUST provide the matching lock value in order for this operation to succeed.");
            }
            finally
            {
                // Release a lock for editing a file.
                commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);
                string identifierForUnlock = isUnlockAndRelockSuccessful ? identifierForUnlockAndRelock : identifierForLock;
                WopiAdapter.UnLock(wopiTargetFileUrl, commonHeaders, identifierForUnlock);
            }
        }

        #endregion 
    }
}