namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WEBSS. Protocol client tries to perform operations associated with activated features. 
    /// </summary>
    [TestClass]
    public class S10_OperationsOnActivatedFeatures : TestSuiteBase
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

        /// <summary> 
        /// This test case verifies the optional behavior of the GetActivatedFeatures operation.
        /// </summary> 
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S10_TC01_GetActivatedFeaturesAboveWSS3()
        {
            Adapter.InitializeService(UserAuthentication.Authenticated);

            Adapter.GetActivatedFeatures();
            if (Common.IsRequirementEnabled(1026, this.Site))
            {
                // Verify MS-WEBSS requirement: MS-WEBSS_R1026
                // When the System Under Test product name is Windows SharePoint Services 3.0 and above, if the server returns an
                //  exception when invoke GetActivatedFeatures operation, then the requirement can be captured.
                this.Site.CaptureRequirement(
                    1026,
                    @"[In Appendix B: Product Behavior] Implementation does support this[GetActivatedFeatures] operation.(<11> Windows SharePoint Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary> 
        /// This test case aims to verify GetActivatedFeatures operation to get a list of activated features on the site and on the parent site collection.
        /// </summary> 
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S10_TC02_GetActivatedFeatures_Succeed()
        {
            Adapter.InitializeService(UserAuthentication.Authenticated);
            string activatedFeatures = Adapter.GetActivatedFeatures();

            // Get the GUIDs of the features on the site.
            string webSiteName = Common.GetConfigurationPropertyValue("webSiteName", this.Site);
            string strFeaturesId = SutAdapter.GetObjectId(webSiteName, FeaturesPosition.SiteFeatures);
            Site.Assert.IsNotNull(strFeaturesId, "This value of the feature id should be non-empty");
            string siteFeatures = strFeaturesId.Remove(0, 1).Replace(
                (char)CONST_CHARS.Blank,
                (char)CONST_CHARS.Comma);

            // Get the GUIDs of the features on the parent site collection.
            strFeaturesId = SutAdapter.GetObjectId(webSiteName, FeaturesPosition.SiteCollectionFeatures);
            Site.Assert.IsNotNull(strFeaturesId, "This value of the feature id should be non-empty");
            string siteCollectionFeatures = strFeaturesId.Remove(0, 1).Replace(
                (char)CONST_CHARS.Blank, (char)CONST_CHARS.Comma);

            char[] commaSeparators = new char[] { (char)CONST_CHARS.Comma };

            string[] activatedFeaturesList = activatedFeatures.Split(
                commaSeparators,
                StringSplitOptions.RemoveEmptyEntries);
            string[] siteFeaturesList = siteFeatures.Split(
                commaSeparators,
                StringSplitOptions.RemoveEmptyEntries);
            string[] siteCollectionFeaturesList = siteCollectionFeatures.Split(
                commaSeparators,
                StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in activatedFeaturesList)
            {
                Guid id;
                Site.Assert.IsTrue(Guid.TryParse(s, out id), "The substring should be a GUID.");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R125
            // COMMENT: If all the substrings split from the result of the operation can be initialized 
            // as a GUID, which means the substrings are GUID; and also the comma-delimited list 
            // of GUIDs of the activated features on the site can be found in the result of the operation, 
            // then the following requirement can be captured.
            bool isVerifiedR125 = false;
            foreach (string s in siteFeaturesList)
            {
                if (activatedFeatures.Contains(s))
                {
                    isVerifiedR125 = true;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR125,
                125,
                @"[In GetActivatedFeatures] The GetActivatedFeaturesResponse element MUST contain a single string, formatted as a comma-delimited list of GUIDs, each identifying an activated feature on the site (2).");

            // Verify MS-WEBSS requirement: MS-WEBSS_R806
            // COMMENT: If all the substrings split from the result of the operation can be initialized 
            // as a GUID, which means the substrings are GUID; and also the comma-delimited list 
            // of GUIDs of the activated features on the site collection can be found in the result of 
            // the operation, then the following requirement can be captured.
            bool isVerifiedR806 = false;
            foreach (string s in siteCollectionFeaturesList)
            {
                if (activatedFeatures.Contains(s))
                {
                    isVerifiedR806 = true;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR806,
                806,
                @"[In GetActivatedFeatures] The GetActivatedFeaturesResponse element MUST contain a single string, formatted as a comma-delimited list of GUIDs, each identifying an activated feature in the site collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R127
            // COMMENT: If all the substrings split from the result of the operation can be initialized 
            // as a GUID, which means the substrings are GUID; and also the comma-delimited list of 
            // GUIDs of the activated features on the site followed by the comma-delimited list of 
            // GUIDs of the activated features on the site collection can be found in the result of the 
            // operation, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR806 && isVerifiedR125,
                127,
                @"[In GetActivatedFeatures] The value of GetActivatedFeaturesResult MUST be a comma-delimited list of GUIDs of features activated on the site (2), followed by a comma-delimited list of GUIDs of features activated on the parent site collection.");

            // Verify MS-WEBSS requirement: MS-WEBSS_R137,MS-WEBSS_R138
            // COMMENT: If all the substrings split from the result of the operation can be initialized 
            // as a GUID, which means the substrings are GUID; and also the comma-delimited list of 
            // GUIDs of the activated features on the site followed by the comma-delimited list of 
            // GUIDs of the activated features on the site collection can be found in the result of the 
            // operation, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR806 && isVerifiedR125,
                137,
                @"[In GetActivatedFeaturesResponse] GetActivatedFeaturesResult: A comma-delimited list of GUIDs, where each GUID is formatted as a UniqueIdentifierWithoutBraces as specified in [MS-WSSCAML] section 2.1.15.");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR806 && isVerifiedR125,
                138,
                @"[In GetActivatedFeaturesResponse] The list MUST include one GUID identifying every feature activated on the site, a tab character followed by one GUID identifying every feature activated in the site collection.");
        }

        /// <summary>
        /// This test case aims to verify the GetActivedFeatures operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S10_TC03_GetActivatedFeatures_Unauthenticated()
        {
            Adapter.InitializeService(UserAuthentication.Unauthenticated);

            try
            {
                Adapter.GetActivatedFeatures();
                Site.Assert.Fail("The expected http status code is not returned for the GetActivatedFeatures operation.");
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                // Verify MS-WEBSS requirement: MS-WEBSS_R1072
                // COMMENT: When the GetActivatedFeatures operation is invoked by unauthenticated 
                // user, if the server return the expected http status code, then the requirement can be 
                // captured.
                HttpWebResponse resp = exp.Response as HttpWebResponse;
                Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    resp.StatusCode,
                    1072,
                    @"[In Message Processing Events and Sequencing Rules] If the user is not authorized to perform the operation[GetActivatedFeatures], the protocol server MUST return an HTTP 401 error as specified in [RFC2616] section 10.4.");
            }
        }

        /// <summary> 
        /// This test case verifies that the GetActivatedFeatures operation aims to verify get a comma-delimited list of GUIDs of activated features on the site or on the parent site collection.
        /// </summary> 
        [TestCategory("MSWEBSS"), TestMethod()]
        public void MSWEBSS_S10_TC04_GetActivatedFeaturesResponseWithCommaDelimitedGUID()
        {
            Adapter.InitializeService(UserAuthentication.Authenticated);

            string activatedFeatures = Adapter.GetActivatedFeatures();

            // Get the GUIDs of the features on the site.
            string webSiteName = Common.GetConfigurationPropertyValue("webSiteName", this.Site);

            string strFeature = SutAdapter.GetObjectId(webSiteName, FeaturesPosition.SiteFeatures);
            Site.Assert.IsNotNull(strFeature, "This value of the feature id should be non-empty");

            // Get the GUIDs of the features on the parent site collection.
            strFeature = SutAdapter.GetObjectId(webSiteName, FeaturesPosition.SiteCollectionFeatures);
            Site.Assert.IsNotNull(strFeature, "This value of the feature id should be non-empty");

            char[] commaSeparators = new char[] { (char)CONST_CHARS.Comma };
            string[] activatedFeaturesList = activatedFeatures.Split(commaSeparators, StringSplitOptions.RemoveEmptyEntries);

            foreach (string s in activatedFeaturesList)
            {
                Guid id;
                Site.Assert.IsTrue(Guid.TryParse(s, out id), "The substring should be a GUID.");
            }

            // Verify MS-WEBSS requirement: MS-WEBSS_R132
            bool isVerifiedR132 = false;
            if (activatedFeatures.Contains(","))
            {
                isVerifiedR132 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR132,
                132,
                @"[In GetActivatedFeaturesSoapOut] It[GetActivatedFeaturesSoapOut] consists of a string consisting of a comma-delimited list of GUIDs, where each GUID identifies a feature activated in the site (2) or the site collection.");
        }
    }
}