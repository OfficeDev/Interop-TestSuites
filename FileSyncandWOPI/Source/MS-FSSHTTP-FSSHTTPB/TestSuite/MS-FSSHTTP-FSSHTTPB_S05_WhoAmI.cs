namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the WhoAmI sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S05_WhoAmI : S05_WhoAmI
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S05_WhoAmI.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S05_WhoAmI.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify the related requirements when the claims-based authentication mode is disabled.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S05_TC01_WhoAmI_ClaimsBased()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Enable the windows-based authentication
                bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchClaimsAuthentication(false);
                this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The windows-based authentication should be enabled successfully.");
                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            WhoAmISubRequestType subRequest = null;
            CellStorageResponse response = null;
            WhoAmISubResponseType subResponse = null;

            while (retryCount > 0)
            {
                // Create a WhoAmI subRequest with all valid parameters
                subRequest = SharedTestSuiteHelper.CreateWhoAmISubRequest(SequenceNumberGenerator.GetCurrentToken());
                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                subResponse = SharedTestSuiteHelper.ExtractSubResponse<WhoAmISubResponseType>(response, 0, 0, this.Site);
                Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "WhoAmI subRequest should be succeed.");

                Regex regex = new Regex(@"^[a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*");
                if (regex.IsMatch(subResponse.SubResponseData.UserLogin))
                {
                    break;
                }

                retryCount--;
                if (retryCount == 0)
                {
                    Site.Assert.Fail("Additional authentication prefix should not exist in UserLogin if claim-based authentication mode is enabled");
                }

                System.Threading.Thread.Sleep(waitTime);
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3087
                Regex regex = new Regex(@"^[a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*");
                bool isVerifiedR3087 = regex.IsMatch(subResponse.SubResponseData.UserLogin);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"The UserLoginType will match pattern ^[a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*, actually its value is {0}",
                    subResponse.SubResponseData.UserLogin);

                this.Site.CaptureRequirementIfIsTrue(
                              isVerifiedR3087,
                              "MS-FSSHTTP",
                              3087,
                              @"[In Appendix B: Product Behavior] If claims-based authentication mode isn't enabled, the implementation does return the UserLogin attribute which apply the following schema:
                              <xs:simpleType name=""UserLoginType"">
                                 <xs:restriction base=""xs:string"">
                                    <xs:pattern value=""^[a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*""/>
                                 </xs:restriction>
                              </xs:simpleType>");
            }

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3086
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Enable the claims-based authentication
                bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchClaimsAuthentication(true);
                this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The claims-based authentication should be enabled successfully.");

                waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
                retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

                while (retryCount > 0)
                {
                    // Send the WhoAmI subRequest to the protocol server
                    response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                    subResponse = SharedTestSuiteHelper.ExtractSubResponse<WhoAmISubResponseType>(response, 0, 0, this.Site);
                    Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "WhoAmI subRequest should be succeed.");

                    Regex r2 = new Regex(@"([a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*)$");
                    Match match = r2.Match(subResponse.SubResponseData.UserLogin);
                    if (r2.IsMatch(subResponse.SubResponseData.UserLogin) && match.Success && match.Index > 0)
                    {
                        break;
                    }

                    retryCount--;
                    if (retryCount == 0)
                    {
                        Site.Assert.Fail("Additional authentication prefix should not exist in UserLogin if claim-based authentication mode is enabled");
                    }

                    System.Threading.Thread.Sleep(waitTime);
                }

                Regex regex2 = new Regex(@"([a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*)$");
                Match m = regex2.Match(subResponse.SubResponseData.UserLogin);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        string.Format("The UserLogin should have prefix when the authentication is claims based, actual the prefix {0}", m.Index > 0 ? "exists" : "does not exist"));

                    this.Site.CaptureRequirementIfIsTrue(
                                  m.Success && m.Index > 0,
                                  "MS-FSSHTTP",
                                  3086,
                                  @"[In Appendix B: Product Behavior] Implementation does add an additional authentication prefix for the UserLogin attribute which is part of the subresponse for a Who Am I subrequest. <28> Section 2.3.2.6: There is an additional authentication prefix if claims-based authentication mode is enabled. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }
                else
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        string.Format("The UserLogin should have prefix when the authentication is claims based, actual the prefix {0}", m.Index > 0 ? "exists" : "does not exist"));

                    this.Site.Assert.IsTrue(
                        m.Success && m.Index > 0,
                        @"[In Appendix B: Product Behavior] Implementation does add an additional authentication prefix for the UserLogin attribute which is part of the subresponse for a Who Am I subrequest. <28> Section 2.3.2.6: There is an additional authentication prefix if claims-based authentication mode is enabled. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S05_TC02_WhoAmI_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a WhoAmI subRequest with all valid parameters.
            WhoAmISubRequestType subRequest = SharedTestSuiteHelper.CreateWhoAmISubRequest(SequenceNumberGenerator.GetCurrentToken());

            CellStorageResponse response = new CellStorageResponse();
            bool isR3006Verified = false;
            try
            {
                // Send a WhoAmI subRequest to the protocol server without specifying URL attribute.
                response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3006Verified = message.Contains("Duplicate attribute");
                isR3006Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3006
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responds two ErrorCode attributes when the URL does not exist.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3006Verified,
                             "MS-FSSHTTP",
                             3006,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3007
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3007,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL does not exist");

                    Site.Assert.IsTrue(
                        isR3006Verified,
                        "[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element doesn't exist, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element is an empty string.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S05_TC03_WhoAmI_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a WhoAmI subRequest with all valid parameters.
            WhoAmISubRequestType subRequest = SharedTestSuiteHelper.CreateWhoAmISubRequest(SequenceNumberGenerator.GetCurrentToken());

            CellStorageResponse response = new CellStorageResponse();
            bool isR3008Verified = false;

            try
            {
                // Send the serverTime subRequest to the protocol server with URL attribute set to en empty string.
                response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException e)
            {
                string message = e.Message;
                isR3008Verified = message.Contains("Duplicate attribute");
                isR3008Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL does not exist");

                    Site.CaptureRequirementIfIsTrue(
                             isR3008Verified,
                             "MS-FSSHTTP",
                             3008,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty string.");

                    Site.Assert.IsTrue(
                        isR3008Verified,
                        "[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. &lt;3&gt; Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// Initialize the shared context based on the specified request file URL, user name, password and domain for the MS-FSSHTTP test purpose.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        protected override void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            SharedContextUtils.InitializeSharedContextForFSSHTTP(userName, password, domain, this.Site);
        }

        /// <summary>
        /// Merge the common configuration and should/may configuration file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        protected override void MergeConfigurationFile(TestTools.ITestSite site)
        {
            ConfigurationFileHelper.MergeConfigurationFile(site);
        }
    }
}