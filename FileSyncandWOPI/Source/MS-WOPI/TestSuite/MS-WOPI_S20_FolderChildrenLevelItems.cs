namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the operations' behaviors on folder children items level whether follow the Open Spec definitions.
    /// </summary>
    [TestClass]
    public class MS_WOPI_S20_FolderChildrenLevelItems : TestSuiteBase
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

        #region Scenario 20

        /// <summary>
        /// This test case is used to verify the EnumerateChildren operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S20_TC01_EnumerateChildren()
        {
            // Get the folder URL.
            string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get folder content URL.
            string wopiFolderContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFolderUrl, WOPISubResourceUrlType.FolderChildrenLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiFolderContentsLevelUrl);

            // Return the contents of a folder on the WOPI server.
            WOPIHttpResponse httpWebResponseForEnumerateChildren = WopiAdapter.EnumerateChildren(wopiFolderContentsLevelUrl, commonHeaders);

            int statusCode = httpWebResponseForEnumerateChildren.StatusCode;

            // Get the json string from the response of EnumerateChildren.
            string jsonStringForEnumerateChildren = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForEnumerateChildren);

            // Convert the json string to object.
            EnumerateChildren enchildren = WOPISerializerHelper.JsonToObject<EnumerateChildren>(jsonStringForEnumerateChildren);
            string fileName = enchildren.Children[0].Name;

            // Verify MS-WOPI requirement: MS-WOPI_R707
            this.Site.CaptureRequirementIfAreEqual<int>(
                          200,
                          statusCode,
                          707,
                          @"[In EnumerateChildren] Status code ""200"" means ""Success"".");

            // The status code is 200 mean success.When response is success the URIs are return.
            this.Site.CaptureRequirement(
                          703,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>/children] Operation ""EnumerateChildren"" is used for ""Returns a set of URIs that provides access to resources in the folder"".");

            // The status code is 200 mean success.When response is success the contexts are return.
            this.Site.CaptureRequirement(
                          704,
                          @"[In EnumerateChildren] The EnumerateChildren method returns the contents of a folder on the WOPI server.");

            // The status code is 200 mean success.When response is success the URI follows the format.
            this.Site.CaptureRequirement(
                          705,
                          @"[In EnumerateChildren] HTTP Verb: GET
URI: HTTP://server/<...>/wopi*/folders/<id>/children?access_token=<token>");

            string subFileUrl = Common.GetConfigurationPropertyValue("UrlOfFileOnSubFolder", this.Site);
            string expectedFileName = TestSuiteHelper.GetFileNameFromFullUrl(subFileUrl);

            // Verify MS-WOPI requirement: MS-WOPI_R713
            bool isEqualToExpectedFileName = expectedFileName.CompareStringValueIgnoreCase(fileName, this.Site);
            this.Site.CaptureRequirementIfIsTrue(
                          isEqualToExpectedFileName,
                          713,
                          @"[In Response Body] Name: The name of the child resource.");

            // Verify MS-WOPI requirement: MS-WOPI_R714
            // The EnumerateChildren request message follow this format and use the id and token return by WOPI server. 
            // If the WOPI server can return the response of EnumerateChildren, capture R714
            this.Site.CaptureRequirement(
                          714,
                          @"[In Response Body] Url: The URI of the child resource of the form http://server/<...>/wopi*/files/<id>?access_token=<token> where id is the WOPI server’s unique id of the resource and token is the token that provides access to the resource.");
        }

        /// <summary>
        /// This test case is used to verify the value of version must change when the file changes and must match the value
        /// which is provided by the "Version" field in the response to CheckFileInfo.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S20_TC02_EnumerateChildren_Version()
        {
            #region Get the WOPI resource URL for visiting file.

            // Get the file URL.
            string fileUrl = Common.GetConfigurationPropertyValue("UrlOfFileOnSubFolder", this.Site);

            // Get the WOPI URL.
            string wopiTargetFileUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(fileUrl, WOPIRootResourceUrlType.FileLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get file content URL.
            string wopiFileContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFileUrl, WOPISubResourceUrlType.FileContentsLevel);

            #endregion 

            #region  Get the WOPI resource URL for visiting folder.

            // Get the folder URL.
            string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get folder content URL.
            string wopiFolderContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFolderUrl, WOPISubResourceUrlType.FolderChildrenLevel);

            #endregion 

            #region Call EnumerateChildren

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiFolderContentsLevelUrl);

            // Return the contents of a folder on the WOPI server.
            WOPIHttpResponse httpWebResponseForEnumerateChildren = WopiAdapter.EnumerateChildren(wopiFolderContentsLevelUrl, commonHeaders);

            // Get the json string from the response of EnumerateChildren.
            string jsonStringForEnumerateChildren = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForEnumerateChildren);

            // Convert the json string to object.
            EnumerateChildren enchildren = WOPISerializerHelper.JsonToObject<EnumerateChildren>(jsonStringForEnumerateChildren);

            string versionOld = enchildren.Children[0].Version;

            #endregion 

            #region Call PutFile

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFileContentsLevelUrl);

            string exceptedUpdateContent = "WOPI PUT file test";
            byte[] bodycontents = Encoding.UTF8.GetBytes(exceptedUpdateContent);
            string identifier = Guid.NewGuid().ToString("N");

            // Update a file on the WOPI server.
            WopiAdapter.PutFile(wopiFileContentsLevelUrl, commonHeaders, null, bodycontents, identifier);

            #endregion 

            #region Call EnumerateChildren

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiFolderContentsLevelUrl);

            // Return the contents of a folder on the WOPI server.
            httpWebResponseForEnumerateChildren = WopiAdapter.EnumerateChildren(wopiFolderContentsLevelUrl, commonHeaders);

            // Get the json string from the response of EnumerateChildren.
            jsonStringForEnumerateChildren = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForEnumerateChildren);

            // Convert the json string to object.
            enchildren = WOPISerializerHelper.JsonToObject<EnumerateChildren>(jsonStringForEnumerateChildren);

            string versionNew = enchildren.Children[0].Version;

            #endregion 

            #region Call CheckFileInfo

            // Get the common header.
            commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFileUrl);

            // Return information about the file 
            WOPIHttpResponse responseOfCheckFileInfo = WopiAdapter.CheckFileInfo(wopiTargetFileUrl, commonHeaders, null);

            // Get the json string from the response of CheckFileInfo.
            string jsonStringForCheckFileInfo = WOPIResponseHelper.ReadHTTPResponseBodyToString(responseOfCheckFileInfo);

            // Convert the json string to object.
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonStringForCheckFileInfo);

            string versionCheckFileInfo = checkFileInfo.Version;

            #endregion 

            // Verify MS-WOPI requirement: MS-WOPI_R716
            this.Site.CaptureRequirementIfAreNotEqual<string>(
                          versionOld,
                          versionNew,
                          716,
                          @"[In Response Body] [Version] This value MUST change when the file changes.");

            // Verify MS-WOPI requirement: MS-WOPI_R934
            this.Site.CaptureRequirementIfAreEqual<string>(
                          versionCheckFileInfo,
                          versionNew,
                          934,
                          @"[In Response Body] [Version] MUST match the value that would be provided by the ""Version"" field in the response to CheckFileInfo (see section 3.3.5.1.1).");
        }

        /// <summary>
        /// This test case is used to verify if the user is unauthorized when call EnumerateChildren operation, server will return 404 error.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S20_TC03_EnumerateChildren_Fail404()
        {
            // Get the folder URL.
            string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get folder content URL.
            string wopiFolderContentsLevelUrl = TokenAndRequestUrlHelper.GetSubResourceUrl(wopiTargetFolderUrl, WOPISubResourceUrlType.FolderChildrenLevel);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiFolderContentsLevelUrl);

            // Remove "Authorization" header from common headers.
            commonHeaders.Remove("Authorization");

            // Return the contents of a folder on the WOPI server with invalid token.
            int statusCode = 0;
            try
            {
                WOPIHttpResponse httpWebResponseForEnumerateChildren = WopiAdapter.EnumerateChildren(wopiFolderContentsLevelUrl, commonHeaders);
                statusCode = httpWebResponseForEnumerateChildren.StatusCode;
            }
            catch (WebException webEx)
            {
                HttpWebResponse errorResponse = this.GetErrorResponseFromWebException(webEx);
                statusCode = this.GetStatusCodeFromHTTPResponse(errorResponse);
            }

            // Verify MS-WOPI requirement: MS-WOPI_R709
            this.Site.CaptureRequirementIfAreEqual<int>(
                          404,
                          statusCode,
                          709,
                          @"[In EnumerateChildren] Status code ""404"" means ""File unknown/User unauthorized"".");
        }

        #endregion 
    }
}