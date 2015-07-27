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
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS_AUTHWSAdapter class.
    /// </summary>
    public interface IMS_AUTHWSAdapter : IAdapter
    {
        /// <summary>
        /// Gets the CookieContainer of web service.
        /// </summary>
        CookieContainer CookieContainer
        {
            get;
        }

        /// <summary>
        /// This operation is used to retrieve the authentication mode that is used by the web application.
        /// </summary>
        /// <returns>An AuthenticationMode value, which specifies the authentication mode for the Login operation.</returns>
        AuthenticationMode Mode();

        /// <summary>
        /// This operation is used to log a user onto the application using the login name and password.
        /// </summary>
        /// <param name="userName">A string containing the login name.</param>
        /// <param name="password">A string containing the password.</param>
        /// <returns>A LoginResult value, which specifies the result of this login operation.</returns>
        LoginResult Login(string userName, string password);

        /// <summary>
        /// This operation is used to switch to the corresponding WebApplication according to AuthenticationMode.
        /// </summary>
        /// <param name="authenicationMode">The current Authentication Mode.</param>
        void SwitchWebApplication(AuthenticationMode authenicationMode);
    }
}