//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to assist MS_OXWSMSGSUTControlAdapter and MS_OXWSSRCHSUTControlAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// Switch the current user to the new user, with the identity of the new role to communicate with server.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="exchangeServiceBinding">An instance of Service Binding</param>
        /// <param name="site">An instance of ITestSite</param>
        public static void SwitchUser(string userName, string password, string domain, ExchangeServiceBinding exchangeServiceBinding, ITestSite site)
        {
            exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, password, domain);

            // Verify the credential of the exchange service binding.
            bool isVerified = false;
            Uri uri = new Uri(Common.GetConfigurationPropertyValue("ServiceUrl", site));
            NetworkCredential credential = exchangeServiceBinding.Credentials.GetCredential(uri, "basic");
            if (credential.Domain == domain && credential.UserName == userName)
            {
                isVerified = true;
            }

            site.Assert.IsTrue(isVerified, "Service binding should be successful");
        }
    }
}
