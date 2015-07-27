//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    /// <summary>
    /// Used to specify the transport type for soap messages.
    /// </summary>
    public enum TransportType
    {
        /// <summary>
        /// Indicate the soap transport over http.
        /// </summary>
        HTTP,

        /// <summary>
        /// Indicate the soap transport over https.
        /// </summary>
        HTTPS
    }

    /// <summary>
    /// This enum indicate the option of authentication information client used.
    /// </summary>
    public enum UserAuthenticationOption
    {
        /// <summary>
        /// Specify Adapter use an authenticated account.
        /// </summary>
        Authenticated,

        /// <summary>
        /// Specify Adapter use an unauthenticated account.
        /// </summary>
        Unauthenticated
    }
}