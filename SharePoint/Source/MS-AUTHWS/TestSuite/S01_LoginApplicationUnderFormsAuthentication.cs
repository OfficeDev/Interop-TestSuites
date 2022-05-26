namespace Microsoft.Protocols.TestSuites.MS_AUTHWS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the login application with Forms authentication.
    /// </summary>
    [TestClass]
    public class S01_LoginApplicationUnderFormsAuthentication : TestSuiteBase
    {
        #region Member Variable Definition

        /// <summary>
        /// An instance of IMS_AUTHWSAdapter, make TestSuite can use IMS_AUTHWSAdapter's function.
        /// </summary>
        private IMS_AUTHWSAdapter authwsAdapter = null;

        /// <summary>
        /// The name of an existing user, who has access to the server.
        /// </summary>
        private string validUserName = null;

        /// <summary>
        /// The password of the user whose account name id specified by the the member variable validUserName.
        /// </summary>
        private string validPassword = null;

        /// <summary>
        /// A string whose value is different from the member variable validPassword.
        /// </summary>
        private string invalidPassword = null;

        /// <summary>
        /// A string whose value is different from the member variable validUserName.
        /// </summary>
        private string invalidUserName = null;

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
        /// This test case is used to verify the Login operation under Forms authentication mode should be successful.
        /// </summary>
        [TestCategory("MSAUTHWS"), TestMethod()]
        public void MSAUTHWS_S01_TC01_VerifyLoginUnderFormsAuthentication()
        {
            // Invoke the Mode operation.
            AuthenticationMode authMode = this.authwsAdapter.Mode();

            bool isVerifyFormsMode = AuthenticationMode.Forms == authMode;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "If the retrieved authentication mode equals to Forms, MS-AUTHWS_R44 can be verified.");

            // If the retrieved authentication mode equals to Forms, MS-AUTHWS_R44 and MS-AUTHWS_R133 can be verified.
            Site.CaptureRequirementIfIsTrue(
              isVerifyFormsMode,
              44,
              @"[In Login] For the operation [Login operation] to succeed, the protocol server MUST use forms authentication and the logon name and password that is provided by the protocol client MUST be valid.");

            Site.CaptureRequirementIfIsTrue(
                isVerifyFormsMode,
                133,
                @"[In Mode] The Mode operation retrieves the authentication mode [Forms] that a Web application uses.");

            int cookieCountBeforeLogin = GetCookieNumber(this.authwsAdapter.CookieContainer);

            // Login application with a valid user name and password, which has access to the server.
            LoginResult loginResult = this.authwsAdapter.Login(this.validUserName, this.validPassword);
            Site.Assert.IsNotNull(loginResult, "Login result is not null");
            Site.Assert.AreEqual<string>(LoginErrorCode.NoError.ToString(), loginResult.ErrorCode.ToString(), "The value of ErrorCode is 'NoError'");

            // If the authentication mode is Forms, and Login application with a valid user name and password succeed verified as above,  MS-AUTHWS_R41 and MS-AUTHWS_R43 can be directly covered.
            Site.CaptureRequirement(
                41,
                @"[In Message Processing Events and Sequencing Rules] The operation ""Login"" logs a user onto an application by using the user’s logon name and password.");

            // Verify MS-AUTHWS requirement: MS-AUTHWS_R43
            Site.CaptureRequirement(
                43,
                @"[In Login] The Login operation logs a user onto a Web application by using the user’s logon name and password.");

            // If the Login application with a valid user name and password succeed, and the returned LoginErrorCode is "NoError" in above steps, MS-AUTHWS_R81 and MS-AUTHWS_R48 can be directly covered.
            Site.CaptureRequirement(
                81,
                @"[In LoginErrorCode] The value of LoginErrorCode is ""NoError""[the authentication mode is Forms], when the Login operation succeeded.");

            // Verify MS-AUTHWS requirement: MS-AUTHWS_R48
            Site.CaptureRequirement(
                48,
                @"[In Login] [If the Login operation succeeds] A redirect to the HTML login form is not performed.");

            int cookieCountAfterLogin = GetCookieNumber(this.authwsAdapter.CookieContainer);
            string lastCookieName = GetCookieName(cookieCountAfterLogin - 1, this.authwsAdapter.CookieContainer);

            bool isCookieCreatedAfterLogin = (cookieCountAfterLogin - cookieCountBeforeLogin == 1) && (!string.IsNullOrEmpty(lastCookieName));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "If the Login with Forms mode succeeded, a ticket for the specified user is created and it is attached to a cookie collection, MS-AUTHWS_R70, MS-AUTHWS_R113, MS-AUTHWS_R72, MS-AUTHWS_R128, MS-AUTHWS_R129 and MS-AUTHWS_R45 can be verified.");

            // If the Login with Forms mode succeeded, a ticket for the specified user is created and it is attached to a cookie collection, MS-AUTHWS_R70, MS-AUTHWS_R113, MS-AUTHWS_R128, MS-AUTHWS_R129 and MS-AUTHWS_R45 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isCookieCreatedAfterLogin,
                70,
                @"[In LoginResult] If a Login WSDL operation succeeded, the name of an authentication cookie [is returned].");

            Site.CaptureRequirementIfIsTrue(
                isCookieCreatedAfterLogin,
                113,
                @"[In AuthenticationMode] If the protocol server authenticates the protocol client, it[protocol server ] issues a cookie to the protocol client [when the AuthenticationMode is ""Forms""] and the protocol client presents that cookie in subsequent requests.");

            Site.CaptureRequirementIfIsTrue(
              isCookieCreatedAfterLogin,
              72,
               @"[In LoginResult] CookieName: A string that specifies the name of the cookie that is used to store the forms authentication ticket.");

            Site.CaptureRequirementIfIsTrue(
                isCookieCreatedAfterLogin,
                128,
                @"[In Login] If the operation [Login operation] succeeds, a ticket for the specified user is created and [it is attached to a cookie collection that is associated with the outgoing response].");

            Site.CaptureRequirementIfIsTrue(
                isCookieCreatedAfterLogin,
                129,
                @"[In Login] If the operation [Login operation] succeeds, a ticket for the specified user is [created and] attached to a cookie collection that is associated with the outgoing response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "If the Login with Forms mode succeeds, and the TimeoutSeconds returned which type is int, MS-AUTHWS_R76 can be verified.");

            // Set R191Enabled to true to verify that implementation does return the element TimeoutSeconds that specifies the number of seconds before the cookie, which is specified in the CookieName element, expires. Set R191Enabled to false to disable this requirement.
            if (Common.IsRequirementEnabled(191, this.Site))
            {
                // If the Login with Forms mode succeeds, and the TimeoutSeconds returned which type is int, MS-AUTHWS_R76 can be verified.
                bool isTimeoutSecondsReturned = loginResult.TimeoutSeconds.GetType() == typeof(int);

                Site.CaptureRequirementIfIsTrue(
                    isTimeoutSecondsReturned,
                    191,
                    @"[In Appendix B: Product Behavior] Implementation does return the element TimeoutSeconds that specifies the number of seconds before the cookie, which is specified in the CookieName element, expires. (The Microsoft SharePoint Foundation 2010, Microsoft SharePoint Foundation 2013 and Microsoft SharePoint Server 2016 follow this behavior.)");
            }

            // Set R126Enabled to true to verify that the default value of the CookieName is "FedAuth" in operation "Login" response. Set R126Enabled to false to disable this requirement.
            if (Common.IsRequirementEnabled(126, this.Site))
            {
                string cookieNameGet = GetCookieName(cookieCountAfterLogin - 1, this.authwsAdapter.CookieContainer);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If the Login with Forms mode succeeds and the returned value of CookieName is 'FedAuth' on Microsoft SharePoint Foundation 2010 and above products, MS-AUTHWS_R126 can be verified.");

                // If the Login with Forms mode succeeds and the returned value of CookieName is 'FedAuth' on Microsoft SharePoint Foundation 2010 and above products, MS-AUTHWS_R126 can be verified.
                Site.CaptureRequirementIfAreEqual<string>(
                    "FedAuth",
                    cookieNameGet,
                    126,
                    @"[In Appendix B: Product Behavior] Implementation does return the default value of the CookieName which is ""FedAuth"".(The Microsoft SharePoint Foundation 2010, Microsoft SharePoint Foundation 2013 and Microsoft SharePoint Server 2016 follow this behavior.)");
            }

            // Set R127Enabled to true to verify that the default value of the CookieName is ".ASPXAUTH" in operation "Login" response. Set R127Enabled to false to disable this requirement.
            if (Common.IsRequirementEnabled(127, this.Site))
            {
                string cookieNameGet = GetCookieName(cookieCountAfterLogin - 1, this.authwsAdapter.CookieContainer);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If the Login with Forms mode succeeds and the returned value of CookieName which is '.ASPXAUTH' on Microsoft SharePoint Server 2007 and Windows SharePoint Services 3.0, MS-AUTHWS_R127 can be verified.");

                // Verify MS-AUTHWS requirement: MS-AUTHWS_R127
                Site.CaptureRequirementIfAreEqual<string>(
                    ".ASPXAUTH",
                    cookieNameGet,
                    127,
                    @"[In Appendix B: Product Behavior] Implementation does return the default value of the CookieName which is "".ASPXAUTH"". <1> Section 3.1.4.1.3.1:  Windows SharePoint Services 3.0 returns the default value of "".ASPXAUTH"".");
            }
        }

        /// <summary>
        /// This test case is used to verify the Login operation under Forms authentication mode is failed with invalid user name.
        /// </summary>
        [TestCategory("MSAUTHWS"), TestMethod()]
        public void MSAUTHWS_S01_TC02_VerifyLoginUnderFormsAuthenticationWithInvalidUserName()
        {
            // Invoke the Mode operation.
            AuthenticationMode authMode = this.authwsAdapter.Mode();
            Site.Assert.AreEqual<AuthenticationMode>(AuthenticationMode.Forms, authMode, "The current authentication mode should be 'Forms', actually the mode is {0}", authMode);

            if (Common.IsRequirementEnabled(830, this.Site))
            {
                try
                {
                    LoginResult loginResult = this.authwsAdapter.Login(this.invalidUserName, this.validPassword);
                }
                catch (System.Web.Services.Protocols.SoapException)
                {
                    Site.Log.Add(LogEntryKind.Debug, "SoapException is returned when the Login operation failed because the logon name is not found by the server, or the password does not match what is stored on the server, MS-AUTHWS_R830 can be verified.");
                }
            }
            else
            {
                // Invoke the Login operation with invalid user name.
                LoginResult loginResult = this.authwsAdapter.Login(this.invalidUserName, this.validPassword);
                Site.Assert.IsNotNull(loginResult, "Login result is not null");
                Site.Assert.IsNull(loginResult.CookieName, "The CookieName is not returned");

                // If the Login operation failed with invalid user name, and the CookieName element is not returned, MS-AUTHWS_74 can be directly verified.
                Site.CaptureRequirement(
                    74,
                    @"[In LoginResult] This element [CookieName element] MUST NOT be present if the Login WSDL operation failed.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If the Login operation failed with invalid user name, and the value of returned ErrorCode is 'PasswordNotMatch', MS-AUTHWS_R84 can be verified.");

                // If the Login operation failed with invalid user name, and the value of returned ErrorCode is 'PasswordNotMatch', MS-AUTHWS_R84 can be verified.
                Site.CaptureRequirementIfAreEqual<LoginErrorCode>(
                    LoginErrorCode.PasswordNotMatch,
                    loginResult.ErrorCode,
                    84,
                    @"[In LoginErrorCode] The value of LoginErrorCode is ""PasswordNotMatch"", when the Login operation failed because the logon name is not found by the server, [or the password does not match what is stored on the server].");
            }
        }

        /// <summary>
        /// This test case is used to verify the Login operation under Forms authentication mode is failed with invalid password.
        /// </summary>
        [TestCategory("MSAUTHWS"), TestMethod()]
        public void MSAUTHWS_S01_TC03_VerifyLoginUnderFormsAuthenticationWithInvalidPassword()
        {
            // Invoke the Mode operation.
            AuthenticationMode authMode = this.authwsAdapter.Mode();
            Site.Assert.AreEqual<AuthenticationMode>(AuthenticationMode.Forms, authMode, "The current authentication mode should be 'Forms', actually the mode is {0}", authMode);

            if (Common.IsRequirementEnabled(830, this.Site))
            {
                try
                {
                    LoginResult loginResult = this.authwsAdapter.Login(this.invalidUserName, this.validPassword);
                }
                catch (System.Web.Services.Protocols.SoapException)
                {
                    Site.Log.Add(LogEntryKind.Debug, "SoapException is returned when the Login operation failed because the logon name is not found by the server, or the password does not match what is stored on the server, MS-AUTHWS_R830 can be verified.");
                }
            }
            else
            {
                // Invoke the Login operation with invalid password.
                LoginResult loginResult = this.authwsAdapter.Login(this.validUserName, this.invalidPassword);
                Site.Assert.IsNotNull(loginResult, "Login result is not null");
                Site.Assert.IsNull(loginResult.CookieName, "The CookieName is not returned");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "If the Login operation failed with invalid password, and the value of returned ErrorCode is 'PasswordNotMatch', MS-AUTHWS_R85 can be verified.");

                // If the Login operation failed with invalid password, and the value of returned ErrorCode is 'PasswordNotMatch', MS-AUTHWS_R85 can be verified.
                Site.CaptureRequirementIfAreEqual<LoginErrorCode>(
                    LoginErrorCode.PasswordNotMatch,
                    loginResult.ErrorCode,
                    85,
                    @"[In LoginErrorCode] The value of LoginErrorCode is ""PasswordNotMatch"", the Login operation failed because the password does not match what is stored on the server.");
            }
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
            this.authwsAdapter.SwitchWebApplication(AuthenticationMode.Forms);

            this.validUserName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            this.validPassword = Common.GetConfigurationPropertyValue("Password", this.Site);
            this.invalidPassword = Common.GenerateInvalidPassword(this.validPassword);
            this.invalidUserName = this.GenerateRandomString(10);
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
