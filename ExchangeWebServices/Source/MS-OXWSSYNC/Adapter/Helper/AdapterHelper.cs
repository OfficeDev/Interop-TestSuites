//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provide common methods for MS_OXWSFOLDSUTControlAdapter and MS_OXWSSYNCSUTControlAdapter
    /// </summary>
    public class AdapterHelper
    {
        #region fields
        /// <summary>
        /// The folders created by case.
        /// </summary>
        private static List<BaseFolderType> createdfolders = new List<BaseFolderType>();
        #endregion

        #region Property
        /// <summary>
        /// Gets or sets the value of field createdfolders.
        /// </summary>
        public static List<BaseFolderType> CreatedFolders
        {
            get { return createdfolders; }
            set { createdfolders = value; }
        }
        #endregion

        /// <summary>
        /// Log on mailbox with specified user account.
        /// </summary>
        /// <param name="name">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="exchangeServiceBinding">An instance of Service Binding</param>
        /// <param name="site">An instance of ITestSite</param>
        /// <returns>If the user logs on mailbox successfully, return true; otherwise, return false.</returns>
        public static bool SwitchUser(string name, string userPassword, string userDomain, ExchangeServiceBinding exchangeServiceBinding, ITestSite site)
        {
            exchangeServiceBinding.Credentials = new NetworkCredential(name, userPassword, userDomain);

            // Verify the credential of the exchange service binding.
            bool isVerified = false;
            Uri uri = new Uri(Common.GetConfigurationPropertyValue("ServiceUrl", site));
            NetworkCredential credential = exchangeServiceBinding.Credentials.GetCredential(uri, "basic");
            if (credential.Domain == userDomain && credential.UserName == name)
            {
                isVerified = true;
            }

            return isVerified;
        }
    }
}
