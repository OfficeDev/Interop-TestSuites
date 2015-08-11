namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with objectId.
    /// </summary>
    [TestClass]
    public class S05_OperationsOnObjectId : TestSuiteBase
    {
        #region Additional test attributes, initialization and clean up

        /// <summary>
        /// Class initialization.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        #region Object Related Cases

        /// <summary>
        /// This test case aims to verify the GetObjectIdFromUrl operation when the objectUrl specifies a list.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S05_TC01_GetObjectIdFromUrl_ValidListUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1032, this.Site), "The test case is executed only when the property 'R1032Enabled' is true.");
            
            GetObjectIdFromUrlResponseGetObjectIdFromUrlResult getObjectIdFromUrlResult = Adapter.GetObjectIdFromUrl(Common.GetConfigurationPropertyValue("GetObjectIdFromUrl_ListUrl", this.Site));

            // If the product is Microsoft SharePoint Foundation 2010 and above, Verify MS-WEBSS requirement: MS-WEBSS_R1032 when invokes the operation "GetObjectIdFromUrl" successfully.
            Site.CaptureRequirement(
                1032,
                @"[In Appendix B: Product Behavior]  Implementation does support this[GetObjectIdFromUrl] operation.(<16>Microsoft SharePoint Foundation 2010 and above follow this behavior.)");

            #region Capture List Related Requirement

            // Verify MS-WEBSS requirement: MS-WEBSS_R328
            string webSiteName = SutAdapter.GetObjectId(Common.GetConfigurationPropertyValue("webSiteName", this.Site), "list");
            Site.Assert.IsNotNull(webSiteName, "This value of the list id should be non-empty");
            string exceptedId = "{" + webSiteName + "}";

            Site.CaptureRequirementIfAreEqual<string>(
                exceptedId,
                getObjectIdFromUrlResult.ObjectId.ListId.ToString(),
                328,
                @"[In GetObjectIdFromUrlResponse] ObjectId.ListId: If the object is a list (1), the value of the attribute MUST be the list identifier.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R1043
            // The type of List Server Template, the values are gotten according to [MS-WSSFO] section 2.2.3.12.
            string listServerTemplateValues = "-1,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,130,140,150,200,201,202,204,207,210,211,212,301,302,303,1100,1200";
            bool isListServerTemplateSpecifiedInMswssfo = listServerTemplateValues.Contains(getObjectIdFromUrlResult.ObjectId.ListServerTemplate.ToString());

            Site.CaptureRequirementIfIsTrue(
                isListServerTemplateSpecifiedInMswssfo,
                1043,
                @"[In GetObjectIdFromUrlResponse] ObjectId.ListServerTemplate: If the object is a list, the value of the attribute MUST be one of the list template types as specified in [MS-WSSFO2] section 2.2.3.12 [the value of the list template type is 100].");

            // The values of List Base Type, the values are gotten according to [MS-WSSFO] section 2.2.3.11.
            string listBaseTypeValues = "0,1,3,4,5";
            bool isListBaseTypeSpecifiedInMswssfo = listBaseTypeValues.Contains(getObjectIdFromUrlResult.ObjectId.ListBaseType.ToString());

            // Verify MS-WEBSS requirement: MS-WEBSS_R1044
            Site.CaptureRequirementIfIsTrue(
                isListBaseTypeSpecifiedInMswssfo,
                1044,
                @"[In GetObjectIdFromUrlResponse] ObjectId.ListBaseType: If the object is a list, the value of the attribute MUST be one of the List Base Types as specified in [MS-WSSFO2] section 2.2.3.11 [the value of the List Base Type is 1].");

            bool isVerifiedR334 = false;
            if (getObjectIdFromUrlResult.ObjectId.ListItem.ToLower().Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedR334 = true;
                Site.Assert.AreEqual("false", getObjectIdFromUrlResult.ObjectId.ListItem.ToLower(), "It is expected to get false, since this is not a list item.");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R334
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR334,
                334,
                @"[In GetObjectIdFromUrlResponse] ObjectId.ListItem: Specifies whether the object is a list item.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R336
            bool isVerifiedR336 = false;
            if (getObjectIdFromUrlResult.ObjectId.ListItemId == null)
            {
                isVerifiedR336 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR336,
                336,
                @"[In GetObjectIdFromUrlResponse] Otherwise[If the object is not a list item], the attribute MUST NOT be present.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R337
            bool isVerifiedR337 = false;
            if (getObjectIdFromUrlResult.ObjectId.File.ToLower().Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedR337 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR337,
                337,
                @"[In GetObjectIdFromUrlResponse] ObjectId.File: Specifies whether the object is a file.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R338
            bool isVerifiedR338 = false;
            if (getObjectIdFromUrlResult.ObjectId.Folder.ToLower().Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                isVerifiedR338 = true;
                Site.Assert.AreEqual("false", getObjectIdFromUrlResult.ObjectId.Folder.ToLower(), "It is expected to get false, since this is not a folder.");
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR338,
                338,
                @"[In GetObjectIdFromUrlResponse] ObjectId.Folder: Specifies whether the object is a folder.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R719
            bool isVerifiedR719 = false;
            string alternateUrl = getObjectIdFromUrlResult.ObjectId.AlternateUrls;
            string[] urlSplit = alternateUrl.Split(',');
            int urlNum = urlSplit.Length;
            string sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            string transportType = Common.GetConfigurationPropertyValue("TransportType", this.Site);
            for (int i = 0; i < urlNum; i++)
            {
                if (urlSplit[i] == transportType.ToLower() + "://" + sutComputerName.ToLower() + "/")
                {
                    isVerifiedR719 = true;
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR719,
                719,
                @"[In GetObjectIdFromUrlResponse] ObjectId. AlternateUrls: Alternate URLs are a comma delimited list of other possible URLs for the object.");

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the GetObjectIdFromUrl operation when the objectUrl specifies a list item.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S05_TC02_GetObjectIdFromUrl_ValidListItemUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1032, this.Site), "The test case is executed only the property 'R1032Enabled' is true.");

            GetObjectIdFromUrlResponseGetObjectIdFromUrlResult getObjectIdFromUrlResult = Adapter.GetObjectIdFromUrl(Common.GetConfigurationPropertyValue("GetObjectIdFromUrl_ListItemUrl", this.Site));

            // If the product is Microsoft SharePoint Foundation 2010 and above, Verify MS-WEBSS requirement: MS-WEBSS_R1032 when invokes the operation "GetObjectIdFromUrl" successfully.
            Site.CaptureRequirement(
                1032,
                @"[In Appendix B: Product Behavior]  Implementation does support this[GetObjectIdFromUrl] operation.(<16>Microsoft SharePoint Foundation 2010 and above follow this behavior.)");

            #region Capture ListItem Related Requirement

            // Verify MS-WEBSS requirement: MS-WEBSS_R313
            // If the specified URL corresponds to an object on the site (2), use that object as input parameter.
            // the server can respond correctly.
            Site.CaptureRequirementIfIsNotNull(
                getObjectIdFromUrlResult,
                313,
                @"[In GetObjectIdFromUrl] If the specified URL corresponds to an object on the site (2), use that object.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R335
            string listItemId = SutAdapter.GetObjectId(Common.GetConfigurationPropertyValue("webSiteName", this.Site), "listItem");
            BaseTestSite.Assert.IsNotNull(listItemId, "This value of the list item should be non-empty");
            Site.CaptureRequirementIfAreEqual<string>(
                listItemId,
                getObjectIdFromUrlResult.ObjectId.ListItemId,
                335,
                @"[In GetObjectIdFromUrlResponse] ObjectId.ListItemId: If the object is a list item, the value of the attribute[ObjectId.ListItemId] MUST be the identifier of the list item.");

            #endregion
        }

        /// <summary>
        /// This test case aims to verify the GetObjectIdFromUrl operation when the objectUrl does not specifies a list or a list item.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S05_TC03_GetObjectIdFromUrl_NoListOrListItemUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1032, this.Site), "The test case is executed only when the property 'R1032Enabled' is true.");

            Adapter.GetObjectIdFromUrl(Common.GetConfigurationPropertyValue("GetObjectIdFromUrl_NoListRelatedUrl", this.Site));
        }

        /// <summary>
        /// This test case aims to verify the GetObjectIdFromUrl operation with invalid objectUrl.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S05_TC04_GetObjectIdFromUrl_InvalidUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1032, this.Site), "The test case is executed only when the property 'R1032Enabled' is true.");

            try
            {
                Adapter.GetObjectIdFromUrl(this.GenerateRandomString(10));
                Site.Assert.Fail("The expected SOAP fault is not returned for the GetObjectIdFromUrl operation.");
            }
            catch (SoapException)
            {
                // Catch the exception, then the following requirements will be captured.
                // Verify MS-WEBSS requirement: MS-WEBSS_R314
                Site.CaptureRequirement(
                    314,
                    @"[In GetObjectIdFromUrl] Otherwise [If the specified URL not corresponds to an object on the site (2), use that object. ], the protocol server MUST return a SOAP exception.");
            }
        }

        /// <summary>
        /// This test case aims to verify the GetObjectIdFromUrl operation when the user is unauthenticated.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S05_TC05_GetObjectIdFromUrl_Unauthenticated()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1032, this.Site), "The test case is executed only when the property 'R1032Enabled' is true.");

            Adapter.InitializeService(UserAuthentication.Unauthenticated);
            try
            {
                Adapter.GetObjectIdFromUrl(Common.GetConfigurationPropertyValue("GetObjectIdFromUrl_ListItemUrl", this.Site));
                Site.Assert.Fail("The expected SOAP fault is not returned for the GetObjectIdFromUrl operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1079
                // COMMENT: When the GetObjectIdFromUrl operation is invoked by unauthenticated user, if 
                // the server return the expected http status code, then the requirement can be captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1079,
                @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetObjectIdFromUrl], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        #endregion
    }
}