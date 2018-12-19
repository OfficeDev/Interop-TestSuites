namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the operations' behaviors on folder level whether follow the Open Spec definitions.
    /// </summary>
    [TestClass]
    public class MS_WOPI_S18_FolderLevelItems : TestSuiteBase
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

        #region Scenario 18

        /// <summary>
        /// This test case is used to verify the CheckFolderInfo operation sequence.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S18_TC01_CheckFolderInfo()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 593, this.Site))
            {
                Site.Assume.Inconclusive(@"The implementation does not support the get the folder access_token and WOPISrc. It is determined using SHOULDMAY PTFConfig property named R593Enabled_MS-WOPI.");
            }
            // Get the folder URL.
            string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

            // Get the WOPI URL.
            string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

            // Get the common header.
            WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFolderUrl);

            // Return information about the folder and permissions that the current user has relative to that file.
            WOPIHttpResponse httpWebResponseForCheckFolderInfo = WopiAdapter.CheckFolderInfo(wopiTargetFolderUrl, commonHeaders, string.Empty);

            // Get the json string from the response of CheckFolderInfo.
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForCheckFolderInfo);

            // Convert the json string to object.
            CheckFolderInfo checkFolderInfo = WOPISerializerHelper.JsonToObject<CheckFolderInfo>(jsonString);

            // Verify MS-WOPI requirement: MS-WOPI_R597
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFolderInfo,
                          597,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>] Operation ""CheckFolderInfo"" is used for ""Returns information about a folder.");

            // Verify MS-WOPI requirement: MS-WOPI_R597
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFolderInfo,
                          598,
                          @"[In CheckFolderInfo] Return information about the folder and permissions that the current user has relative to that file.");

            // The URI in "CheckFolderInfo" WOPI request follow the "HTTP://server/<...>/wopi*/folders/<id>?access_token=<token>" pattern, if the operation execute successfully, capture R599
            // Verify MS-WOPI requirement: MS-WOPI_R599
            this.Site.CaptureRequirement(
                          599,
                          @"[In CheckFolderInfo] HTTP Verb: GET
                          URI: HTTP://server/<...>/wopi*/folders/<id>?access_token=<token>");

            if (!string.IsNullOrEmpty(checkFolderInfo.UserFriendlyName))
            {
                // Verify MS-WOPI requirement: MS-WOPI_R645
                bool isVerifiedR645 = checkFolderInfo.UserFriendlyName.IndexOf(Common.GetConfigurationPropertyValue("UserName", this.Site), StringComparison.OrdinalIgnoreCase) >= 0 ||
                    checkFolderInfo.UserFriendlyName.IndexOf(Common.GetConfigurationPropertyValue("UserFriendlyName", this.Site), StringComparison.OrdinalIgnoreCase) >= 0;

                this.Site.CaptureRequirementIfIsTrue(
                              isVerifiedR645,
                              645,
                              @"[In Response Body] UserFriendlyName: A string that is the name of the user.");
            }
        }

        /// <summary>
        /// This test case is used to verify the value of the HostAuthenticationId from the response of
        /// CheckFolderInfo operation is unique.
        /// </summary>
        [TestCategory("MSWOPI"), TestMethod()]
        public void MSWOPI_S18_TC02_CheckFolderInfo_HostAuthenticationId()
        {
            if (!Common.IsRequirementEnabled("MS-WOPI", 593, this.Site))
            {
                Site.Assume.Inconclusive(@"The implementation does not support the get the folder access_token and WOPISrc. It is determined using SHOULDMAY PTFConfig property named R593Enabled_MS-WOPI.");
            }
            {
                // Get the folder URL.
                string folderFullUrl = Common.GetConfigurationPropertyValue("SubFolderUrl", this.Site);

                // Get the WOPI URL.
                string wopiTargetFolderUrl = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(folderFullUrl, WOPIRootResourceUrlType.FolderLevel, TokenAndRequestUrlHelper.DefaultUserName, TokenAndRequestUrlHelper.DefaultPassword, TokenAndRequestUrlHelper.DefaultDomain);

                // Get the common header.
                WebHeaderCollection commonHeaders = HeadersHelper.GetCommonHeaders(wopiTargetFolderUrl);

                // Return information about the folder and permissions that the current user has relative to that file.
                WOPIHttpResponse httpWebResponseForCheckFolderInfo = WopiAdapter.CheckFolderInfo(wopiTargetFolderUrl, commonHeaders, string.Empty);

                string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForCheckFolderInfo);

                // Convert the json string to object.
                CheckFolderInfo checkFolderInfo = WOPISerializerHelper.JsonToObject<CheckFolderInfo>(jsonString);

                // Get the WOPI URL.
                string wopiTargetFolderUrlOther = WopiSutManageCodeControlAdapter.GetWOPIRootResourceUrl(
                    folderFullUrl,
                    WOPIRootResourceUrlType.FileLevel,
                    Common.GetConfigurationPropertyValue("UserName1", this.Site),
                    Common.GetConfigurationPropertyValue("Password1", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site));

                // Get the common header.
                WebHeaderCollection commonHeadersOther = HeadersHelper.GetCommonHeaders(wopiTargetFolderUrlOther);

                // Return information about the folder and permissions that the current user has relative to that file.
                WOPIHttpResponse httpWebResponseForCheckFolderInfoOther = WopiAdapter.CheckFolderInfo(wopiTargetFolderUrlOther, commonHeadersOther, string.Empty);

                jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(httpWebResponseForCheckFolderInfoOther);

                // Convert the json string to object.
                CheckFolderInfo checkFolderInfoOther = WOPISerializerHelper.JsonToObject<CheckFolderInfo>(jsonString);

                // Verify MS-WOPI requirement: MS-WOPI_R742
                this.Site.CaptureRequirementIfAreNotEqual(
                              checkFolderInfo.HostAuthenticationId,
                              checkFolderInfoOther.HostAuthenticationId,
                              742,
                              @"[In Response Body] HostAuthenticationId: A string that is used by the WOPI server to uniquely for a sample of N (default N=2) identify the users.");

                // Verify requirement MS-WOPI_R59
                bool isR59Satisfied = checkFolderInfo.HostAuthenticationId != checkFolderInfoOther.HostAuthenticationId;

                if (Convert.ToBoolean(Common.IsRequirementEnabled("MS-WOPI", 59, this.Site)))
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R59
                    this.Site.CaptureRequirementIfIsTrue(
                                  isR59Satisfied,
                                  59,
                                  @"[In Common URI Parameters] Implementation does support the token is scoped to a specific user and set of resources.(Microsoft SharePoint Foundation 2013 and above follow this behavior)");
                }
            }
        }
        #endregion
    }
}