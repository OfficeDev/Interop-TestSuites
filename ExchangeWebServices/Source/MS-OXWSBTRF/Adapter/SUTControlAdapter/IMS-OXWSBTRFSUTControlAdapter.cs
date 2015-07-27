//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSBTRF SUT control adapter interface.
    /// </summary>
    public interface IMS_OXWSBTRFSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Log on to a mailbox with a specified user account and delete all the subfolders from the specified folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be cleaned up.</param>
        /// <param name="destFolderName">The name of the destination folder which will be deleted.</param>
        /// <returns>If the folder is cleaned up successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and delete the specified folder (destFolderName) from the (folderName) folder." +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool CleanupFolder(string userName, string userPassword, string userDomain, string folderName, string destFolderName);
    }
}