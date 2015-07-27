//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The SUT control adapter of MS-OXWSSRCH.
    /// It includes methods FindItem() and IsItemAvailableAfterMoveOrDelete() which can be implemented with operations defined in MS-OXWSSRCH.
    /// </summary>
    public interface IMS_OXWSSRCHSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// The operation searches the specified user's mailbox and returns the result whether one or more valid items are found.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="domain">Domain of the user.</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the operation succeeds, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and find the item in the (folderName) folder whose referenced field URI (field) is (value).\n" +
            "If the item is found, enter \"true\"; " +
            "otherwise, enter \"false\".")]
        bool FindItem(string userName, string password, string domain, string folderName, string value, string field);

        /// <summary>
        /// The operation searches the specified user's mailbox and returns the result whether the valid items are found after the MoveItem or DeleteItem operation completed.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="domain">Domain of the user.</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the item existed in the specific folder, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and check whether the item still exists in the (folderName) folder whose referenced field URI (field) is (value).\n" +
            "If the item still exists in the specific folder, enter \"true\"; " +
            "otherwise, enter \"false\".")]
        bool IsItemAvailableAfterMoveOrDelete(string userName, string password, string domain, string folderName, string value, string field);
    }
}