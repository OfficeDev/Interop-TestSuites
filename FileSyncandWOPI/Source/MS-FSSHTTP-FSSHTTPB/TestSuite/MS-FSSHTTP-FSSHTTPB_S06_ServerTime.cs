namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the ServerTime sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S06_ServerTime : S06_ServerTime
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S06_ServerTime.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S06_ServerTime.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element is an empty string.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S06_TC01_ServerTime_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a serverTime subRequest with all valid parameters.
            ServerTimeSubRequestType subRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());

            CellStorageResponse response = new CellStorageResponse();
            bool isR3008Verified = false;
            try
            {
                // Send a ServerTime subRequest to the protocol server without specifying URL attribute.
                response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3008Verified = message.Contains("Duplicate attribute");
                isR3008Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3008
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3008Verified,
                             "MS-FSSHTTP",
                             3008,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. <3> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
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
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S06_TC02_ServerTime_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a serverTime subRequest with all valid parameters.
            ServerTimeSubRequestType subRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());

            CellStorageResponse response = new CellStorageResponse();
            bool isR3006Verified = false;
            try
            {
                // Send a ServerTime subRequest to the protocol server without specifying URL attribute.
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
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

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
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.Assert.IsTrue(
                        isR3006Verified,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <3> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
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