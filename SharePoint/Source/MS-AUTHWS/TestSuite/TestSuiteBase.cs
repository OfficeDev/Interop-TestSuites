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
    using System;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Contain test cases designed to test [MS_AUTHWS] protocol.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion Test Suite Initialization
        
        #region Test Case Initialization

        /// <summary>
        /// Initialize the test.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
        }

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #endregion Test Case Initialization

        #region Protected Methods that used by child test suite class

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        protected string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// This method is used to get count of the cookies exist on the server.
        /// </summary>
        /// <param name="cookieContainer">The cookie container which contains the specified cookie.</param>
        /// <returns>The cookie number exists on the server.</returns>
        protected int GetCookieNumber(CookieContainer cookieContainer)
        {
            CookieCollection cookieCollection = cookieContainer.GetCookies(new Uri(this.GetFormsAuthenticationServiceUrl()));
            if (cookieCollection != null)
            {
                return cookieCollection.Count;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// This method is used to get the specified cookie name exists on the server.
        /// </summary>
        /// <param name="cookieIndex">The cookie index which name to be returned.</param>
        /// <param name="cookieContainer">The cookie container which contains the specified cookie.</param>
        /// <returns>The specified cookie name exists on the server.</returns>
        protected string GetCookieName(int cookieIndex, CookieContainer cookieContainer)
        {
            string cookieName = string.Empty;
            CookieCollection cookieCollection = cookieContainer.GetCookies(new Uri(this.GetFormsAuthenticationServiceUrl()));
            
            if (cookieCollection != null)
            {
                if (cookieCollection.Count != 0)
                {
                    Cookie[] cookies = new Cookie[cookieCollection.Count];
                    cookieCollection.CopyTo(cookies, 0);
                    cookieName = cookies[cookieIndex].Name;
                }
            }

            return cookieName;
        }

        /// <summary>
        /// This method is used to get the Forms Authentication Service Url on the server.
        /// </summary>
        /// <returns>The Forms Authentication Service Url on the server.</returns>
        private string GetFormsAuthenticationServiceUrl()
        {
            string url = string.Empty;

            if (Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site) == TransportProtocol.HTTP)
            {
                url = Common.GetConfigurationPropertyValue("FormsAuthenticationUrlForHTTP", this.Site);
            }
            else
            {
                url = Common.GetConfigurationPropertyValue("FormsAuthenticationUrlForHTTPS", this.Site);
            }

            return url;
        }

    #endregion Protected Methods that used by child test suite class
    }
}