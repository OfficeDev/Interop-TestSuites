//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    ///  A form of encoding of an internal identifier that makes it globally unique 
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct GID
    {
        /// <summary>
        /// A value that represents a namespace for IDs. 
        /// </summary>
        public Guid DatabaseGuid;

        /// <summary>
        /// An auto-incrementing 6-byte value.
        /// </summary>
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 6)]
        public byte[] GlobalCounter;
    }
}