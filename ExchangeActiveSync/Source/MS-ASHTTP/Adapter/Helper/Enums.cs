//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    /// <summary>
    /// The fields in request line and request headers which need to be specially configured besides command name and command parameters.
    /// </summary>
    public enum HTTPPOSTRequestPrefixField
    {
        /// <summary>
        /// The header encoding type.
        /// </summary>
        QueryValueType,

        /// <summary>
        /// The version of ActiveSync protocol.
        /// </summary>
        ActiveSyncProtocolVersion,

        /// <summary>
        /// The AcceptEncoding request header.
        /// </summary>
        AcceptEncoding,

        /// <summary>
        /// The User-Agent request header.
        /// </summary>
        UserAgent,

        /// <summary>
        /// The MS-ASAcceptMultiPart request header.
        /// </summary>
        AcceptMultiPart,

        /// <summary>
        /// The X-MS-PolicyKey request header.
        /// </summary>
        PolicyKey,

        /// <summary>
        /// The prefix of request URI.
        /// </summary>
        PrefixOfURI,

        /// <summary>
        /// The computer name of SUT in the request URI.
        /// </summary>
        Host,

        /// <summary>
        /// The user name of the Authorization request header.
        /// </summary>
        UserName,

        /// <summary>
        /// The password of Authorization request header.
        /// </summary>
        Password,

        /// <summary>
        /// The domain of Authorization request header.
        /// </summary>
        Domain
    }
}