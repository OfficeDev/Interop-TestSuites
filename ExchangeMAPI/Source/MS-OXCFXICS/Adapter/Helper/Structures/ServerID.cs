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
    using System.Runtime.InteropServices;

    /// <summary>
    /// The PtypServerId Type
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct ServerID 
    {
        /// <summary>
        /// 0x01 indicates the remaining bytes conform to this structure; 
        /// 0x00 indicates this is a client-defined value, 
        /// and has whatever size and structure the client has defined.
        /// </summary>
        public byte Ours;

        /// <summary>
        /// A FID, identifying a folder.
        /// </summary>
        public ulong FID;

        /// <summary>
        /// A MID,identifying a message in the folder identified by folder ID.
        /// If the object is a folder, then this field MUST be all zeros.
        /// </summary>
        public ulong MID;

        /// <summary>
        /// A 32-bit unsigned instance number within an array 
        /// of ServerIds to compare against
        /// </summary>
        public uint Instance;
    }
}