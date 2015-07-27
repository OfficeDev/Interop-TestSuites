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
    /// <summary>
    /// The FolderProperties
    /// </summary>
    public enum FolderProperties : uint
    {
        /// <summary>
        /// The PidTagContainerContents
        /// </summary>
        PidTagContainerContents = 0x360f000d,

        /// <summary>
        /// The PidTagFolderAssociatedContents
        /// </summary>
        PidTagFolderAssociatedContents = 0x3610000d,

        /// <summary>
        /// The PidTagContainerHierarchy
        /// </summary>
        PidTagContainerHierarchy = 0x360E000d
    }
}