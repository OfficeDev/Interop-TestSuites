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
    /// Represents the type of FastTransfer stream.
    /// </summary>
    public enum FastTransferStreamType
    {
        /// <summary>
        /// The contentsSync.
        /// </summary>
        contentsSync = 0,

        /// <summary>
        /// The hierarchySync.
        /// </summary>
        hierarchySync = 1,

        /// <summary>
        /// The current state.
        /// </summary>
        state = 2,

        /// <summary>
        /// The folderContent.
        /// </summary>
        folderContent = 3,

        /// <summary>
        /// The MessageContent.
        /// </summary>
        MessageContent = 4,

        /// <summary>
        /// The attachmentContent.
        /// </summary>
        attachmentContent = 5,

        /// <summary>
        /// The MessageList.
        /// </summary>
        MessageList = 6,

        /// <summary>
        /// The TopFolder.
        /// </summary>
        TopFolder = 7
    }
}