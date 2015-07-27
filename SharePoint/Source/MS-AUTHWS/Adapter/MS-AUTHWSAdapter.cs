//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_AUTHWS
{
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS_AUTHWS.
    /// </summary>
    public partial class MS_AUTHWSAdapter : ManagedAdapterBase, IMS_AUTHWSAdapter
    {
        #region Member Variables

        /// <summary>
        /// An instance of AuthenticationSoap class, which represents the authentication web service stub.
        /// </summary>
        private AuthenticationSoap authwsServiceStub;

        /// <summary>
        /// Gets the CookieContainer of web service.
        /// </summary>
        public CookieContainer CookieContainer
        {
            get
            {
                return this.authwsServiceStub.CookieContainer;
            }
        }

        #endregion

        #region Initialize TestSuite

        /// <summary>
        /// This method overrides Initialize in the base class to do initialization works.
        /// </summary>
        /// <param name="testSite">The test site which delivers information to initialize the adapter.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            this.authwsServiceStub = Proxy.CreateProxy<AuthenticationSoap>(this.Site);

            Site.DefaultProtocolDocShortName = "MS-AUTHWS";

            // Merge the common configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);

            this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("WindowsAuthenticationUrlForHTTP", this.Site);

            this.authwsServiceStub.SoapVersion = Common.GetConfigurationPropertyValue<SoapProtocolVersion>("SoapVersion", this.Site);

            this.authwsServiceStub.Credentials = new NetworkCredential(Common.GetConfigurationPropertyValue("UserName", this.Site), Common.GetConfigurationPropertyValue("Password", this.Site), Common.GetConfigurationPropertyValue("Domain", this.Site));
            this.authwsServiceStub.CookieContainer = new CookieContainer();
        }

        /// <summary>
        /// The method overrides Reset in the base class to reset the CookieContainer property of the service stub class.
        /// </summary>
        public override void Reset()
        {
            this.authwsServiceStub.CookieContainer = new CookieContainer();
            base.Reset();
        }

        #endregion Initialize TestSuite

        #region MS_AUTHWSAdapter Members

        /// <summary>
        /// This operation invokes the Mode operation in the service stub authwsServiceStub 
        /// to get the authentication mode that is used by the server
        /// and verifies the message structure and xml schema of the response.
        /// </summary>
        /// <returns>An AuthenticationMode value, which specifies the authentication mode for the Login operation.</returns>
        public AuthenticationMode Mode()
        {
            AuthenticationMode authenticationMode = this.authwsServiceStub.Mode();

            this.ValidateModeResponse();
            this.CaptureTransportRelatedRequirements();

            return authenticationMode;
        }

        /// <summary>
        /// This operation is used to invoke the Login operation of the service stub authwsServiceStub
        /// to log a user onto the server
        /// and verifies the messages structure and xml schema of the response.
        /// </summary>
        /// <param name="userName">A string containing the login name.</param>
        /// <param name="password">A string containing the password.</param>
        /// <returns>A LoginResult value, which specifies the result of this login operation.</returns>
        public LoginResult Login(string userName, string password)
        {
            LoginResult loginResult = this.authwsServiceStub.Login(userName, password);

            this.ValidateLoginResponse(loginResult);
            this.CaptureTransportRelatedRequirements();

            return loginResult;
        }

        /// <summary>
        /// This operation is used to switch to the corresponding WebApplication according to AuthenticationMode.
        /// </summary>
        /// <param name="authenticationMode">The current Authentication Mode.</param>
        public void SwitchWebApplication(AuthenticationMode authenticationMode)
        {
            TransportProtocol currenTransportType = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);

            if (currenTransportType == TransportProtocol.HTTP)
            {
                switch (authenticationMode)
                {
                    case AuthenticationMode.Forms:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("FormsAuthenticationUrlForHTTP", this.Site);
                        break;
                    case AuthenticationMode.None:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("NoneAuthenticationUrlForHTTP", this.Site);
                        break;
                    case AuthenticationMode.Passport:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("PassportAuthenticationUrlForHTTP", this.Site);
                        break;
                    case AuthenticationMode.Windows:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("WindowsAuthenticationUrlForHTTP", this.Site);
                        break;
                }
            }
            else if (currenTransportType == TransportProtocol.HTTPS)
            {
                // When request Url include HTTPS prefix, avoid closing base connection.
                // Local client will accept all certificates after executing this function. 
                Common.AcceptServerCertificate();

                switch (authenticationMode)
                {
                    case AuthenticationMode.Forms:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("FormsAuthenticationUrlForHTTPS", this.Site);
                        break;
                    case AuthenticationMode.None:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("NoneAuthenticationUrlForHTTPS", this.Site);
                        break;
                    case AuthenticationMode.Passport:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("PassportAuthenticationUrlForHTTPS", this.Site);
                        break;
                    case AuthenticationMode.Windows:
                        this.authwsServiceStub.Url = Common.GetConfigurationPropertyValue("WindowsAuthenticationUrlForHTTPS", this.Site);
                        break;
                }
            }
            else
            {
                Site.Assume.Fail("The property 'TransportType' value at the ptfconfig file must be HTTP or HTTPS.");
            }
        }

        #endregion MS_AUTHWSAdapter Members
    }
}