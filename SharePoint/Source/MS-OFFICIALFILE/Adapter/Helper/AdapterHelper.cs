//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.Protocols.TestTools;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;

    /// <summary>
    /// Used to read configuration property from ptfconfig and capture requirements.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// A object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        private static ITestSite testSite;

        /// <summary>
        /// Gets or sets an object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        public static ITestSite Site
        {
            get { return AdapterHelper.testSite; }
            set { AdapterHelper.testSite = value; }
        }

        /// <summary>
        ///  Initialize this helper class with ITestSite used to read configuration property from ptfconfig and capture requirements.
        /// </summary>
        /// <param name="site">A object provides logging, assertions, and SUT adapters for test code onto its execution context.</param>
        public static void Initialize(ITestSite site)
        {
            Site = site;
        }

        /// <summary>
        /// Gets properties from ptfconfig.
        /// </summary>
        /// <param name="propertyName">Property name.</param>
        /// <returns>Property value.</returns>
        public static string GetProperty(string propertyName)
        {
            if (Common.Common.GetConfigurationPropertyValue(propertyName, Site) == null)
            {
                Site.Assume.Fail("Property {0} was not found in the configuration", propertyName);
            }

            return Common.Common.GetConfigurationPropertyValue(propertyName, Site);
        }

        /// <summary>
        /// Case request Url include HTTPS prefix, use this function to avoid closing base connection.
        /// Local client will accept all certificate after execute this function. 
        /// </summary>
        public static void AcceptAllCertificate()
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
        }

        /// <summary>
        /// Set the security credential for Web service client authentication.
        /// </summary>
        /// <param name="userName">This web service's user name.</param>
        /// <param name="password">This web service's password.</param>
        /// <param name="domainName">This web service's domain name.</param>
        /// <returns>A security credential.</returns>
        public static ICredentials ConfigureCredential(string userName, string password, string domainName)
        {
            if (userName == null)
            {
                return CredentialCache.DefaultCredentials;
            }
            else if (password == null)
            {
                return new NetworkCredential(userName, domainName);
            }
            else
            {
                return new NetworkCredential(userName, password, domainName);
            }
        }

        /// <summary>
        /// Convert the binary input into Base64 output.
        /// </summary>
        /// <param name="toEncode">A string indicates the content of file submitted.</param>
        /// <returns>A byte array indicates Base64 output.</returns>
        public static byte[] EncodeToBase64(string toEncode)
        {
            byte[] toEncodeAsBytes

                  = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);

            string stringValue

                  = System.Convert.ToBase64String(toEncodeAsBytes);

            return System.Text.ASCIIEncoding.ASCII.GetBytes(stringValue);
        }

        /// <summary>
        /// Verifies the remote Secure Sockets Layer (SSL) certificate used for authentication.
        /// In our adapter, we make this method always return true, make client can communicate with server under HTTPS without a certification. 
        /// </summary>
        /// <param name="sender">An object that contains state information for this validation.</param>
        /// <param name="certificate">The certificate used to authenticate the remote party.</param>
        /// <param name="chain">The chain of certificate authorities associated with the remote certificate.</param>
        /// <param name="sslPolicyErrors">One or more errors associated with the remote certificate.</param>
        /// <returns>A Boolean value that determines whether the specified certificate is accepted for authentication.</returns>
        private static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
    }
}