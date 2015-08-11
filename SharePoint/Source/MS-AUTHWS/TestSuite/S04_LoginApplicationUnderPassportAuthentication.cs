namespace Microsoft.Protocols.TestSuites.MS_AUTHWS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the Login operation under Passport authentication mode.
    /// </summary>
    [TestClass]
    public class S04_LoginApplicationUnderPassportAuthentication : TestSuiteBase
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
        /// This test case is used to verify the Mode and Login operations when the SUT's authentication mode is Passport.
        /// </summary>
        [TestCategory("MSAUTHWS"), TestMethod()]
        public void MSAUTHWS_S04_TC01_VerifyLoginUnderPassportAuthentication()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(132, this.Site), "This case runs only when the requirement 132 is enabled.");

            // Invoke Mode operation.
            AuthenticationMode authMode = this.authwsAdapter.Mode();

            bool isVerifyPassportMode = AuthenticationMode.Passport == authMode;
            Site.Assume.IsTrue(isVerifyPassportMode, string.Format("The expected result of Mode is Passport, the actual result is{0}", authMode.ToString()));

            // Invoke Login operation.
            LoginResult loginResult = this.authwsAdapter.Login(this.validUserName, this.validPassword);
            Site.Assert.IsNotNull(loginResult, "Login result is not null");
            Site.Assert.IsNull(loginResult.CookieName, "The cookie name is null");
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
            this.authwsAdapter.SwitchWebApplication(AuthenticationMode.Passport);

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