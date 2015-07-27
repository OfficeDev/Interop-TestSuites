//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Error code returned by RPC calls.
    /// </summary>
    public static class OxcRpcErrorCode
    {
        /// <summary>
        /// The operation succeeded.
        /// </summary>
        public const uint ECNone = 0x00000000;

        /// <summary>
        /// A badly formatted RPC buffer was detected.
        /// </summary>
        public const uint ECRpcFormat = 0x000004B6;

        /// <summary>
        /// (0x0000047D)Error code returned when response payload bigger than the maximum of pcbOut.
        /// </summary>
        public const uint ECResponseTooBig = 0x0000047D;

        /// <summary>
        /// (0x0000047D)Error code returned when a buffer passed to this function is not big enough.
        /// </summary>
        public const uint ECBufferTooSmall = 0x0000047D;
    }
}