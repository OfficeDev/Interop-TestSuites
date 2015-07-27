//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    /// <summary>
    /// Address families.
    /// </summary>
    public enum AddressFamily : ushort
    {
        /// <summary>
        /// Internetwork: UDP, TCP, etc.
        /// </summary>
        AF_INET = 2,

        /// <summary>
        /// Internetwork Version 6
        /// </summary>
        AF_INET6 = 23,

        /// <summary>
        /// Invalid address family
        /// </summary>
        Invalid = 0xff
    }
}