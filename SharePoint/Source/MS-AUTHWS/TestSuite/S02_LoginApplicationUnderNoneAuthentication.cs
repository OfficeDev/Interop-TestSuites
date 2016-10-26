namespace Microsoft.Protocols.TestSuites.MS_AUTHWS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    ///  This scenario is designed to test the logon application with None authentication.
    /// </summary>
    [TestClass]
    public class S02_LoginApplicationUnderNoneAuthentication : TestSuiteBase
    {
        #region Member Variable Definition

        /// <summary>
        /// An instance of IMS_AUTHWSAdapter, make TestSuite can use IAUTHWSAdapter's function.
        /// </summary>
        private IMS_AUTHWSAdapter authwsAdapter = null;

        /// <summary>
        /// The name of an existing user, who has access to the server.
        /// </summary>
        private string validUserName = null;

        /// <summary>
        /// The password of the user whose account name id specified by the member variable validUserName.
        /// </summary>
        private string validPassword = null;

        #endregion

        #region Test Suite Initialize and Cleanup

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is used to verify the Login operation under None authentication is failed.
        /// </summary>
        [TestCategory("MSAUTHWS"), TestMethod()]
        public void MSAUTHWS_S02_TC01_VerifyLoginUnderNoneAuthentication()
        {
            // Invoke the Mode operation.
            AuthenticationMode authMode = this.authwsAdapter.Mode();

            // Set R193Enabled to true to verify that implementation does not use authentication if the AuthenticationMode is "None". Set R193Enabled to false to disable this requirement.
            if (Common.IsRequirementEnabled(193, this.Site))
            {
                bool isVerifyNoneMode = AuthenticationMode.None == authMode;

                // If the retrieved authentication mode equals to None, MS-AUTHWS_134 can be verified. 
                Site.CaptureRequirementIfIsTrue(
                    isVerifyNoneMode,
                    134,
                    @"[In Mode] The Mode operation retrieves the authentication mode [None] that a Web application uses.");

                // If the retrieved authentication mode equals to None, MS-AUTHWS_193 can be verified. 
                Site.CaptureRequirementIfIsTrue(
                    isVerifyNoneMode,
                    193,
                    @"[In Appendix B: Product Behavior] Implementation does not use authentication if the AuthenticationMode is ""None"". (Windows SharePoint Services 3.0, Microsoft SharePoint Foundation 2013 and Microsoft SharePoint Server 2016 follow this behavior.)");
            }

            // Invoke the Login operation.
            LoginResult loginResult = this.authwsAdapter.Login(this.validUserName, this.validPassword);
            Site.Assert.IsNotNull(loginResult, "Login result is not null");
            Site.Assert.IsNull(loginResult.CookieName, "The cookie name is not returned");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "If the Login operation failed, and the value of returned ErrorCode is 'NotInFormsAuthenticationMode', MS-AUTHWS_R130 can be verified.");

            // If the Login operation failed, and the value of returned ErrorCode is 'NotInFormsAuthenticationMode', MS-AUTHWS_R130 can be verified.
            Site.CaptureRequirementIfAreEqual<LoginErrorCode>(
                LoginErrorCode.NotInFormsAuthenticationMode,
                loginResult.ErrorCode,
                130,
                @"[In LoginErrorCode] The value of LoginErrorCode is ""NotInFormsAuthenticationMode"", when the Login operation failed because the authentication mode is set to None authentication.");
        }
        #endregion Test Cases

        #region Test Method Initialize and Teardown

        /// <summary>
        /// Overrides OfficeProtocolTestClass's TestInitialize(), to initialize the instance of IMS_AUTHWSAdapter and get some configuration properties.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.authwsAdapter = this.Site.GetAdapter<IMS_AUTHWSAdapter>();
            this.authwsAdapter.SwitchWebApplication(AuthenticationMode.None);

            this.validUserName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            this.validPassword = Common.GetConfigurationPropertyValue("Password", this.Site);
        }

        /// <summary>
        /// Override OfficeProtocolTestClass's TestCleanup(), set server's authentication mode back to Windows after each test case.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #endregion
    }
}