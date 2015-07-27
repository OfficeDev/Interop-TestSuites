//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ASCMD SUT control adapter interface
    /// </summary>
    public interface IMS_ASCMDSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Deletes a user's ActiveSync device
        /// </summary>
        /// <param name="serverComputerName">The computer name of the server.</param>
        /// <param name="userName">The name of the user, who is in administrators group of server, used to communicate with the server.</param>
        /// <param name="userPassword">The password of the user used to communicate with server.</param>
        /// <param name="userDomain">The domain of the user used to communicate with server.</param>
        /// <returns>If user's ActiveSync device is deleted, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to the server (serverComputerName) with the specified user account (userName, userPassword, userDomain), " +
            "and delete the specified user's all ActiveSync devices. " +
            "If the user's ActiveSync devices are deleted successfully, enter \"true\"; otherwise, enter \"false\".")]
        bool DeleteDevice(string serverComputerName, string userName, string userPassword, string userDomain);

        /// <summary>
        /// Gets the value of SUT's AccessRights property.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the server.</param>
        /// <param name="userName">The name of the user, who is in administrators group of server, used to communicate with the server.</param>
        /// <param name="userPassword">The password of the user used to communicate with server.</param>
        /// <param name="userDomain">The domain of the user used to communicate with server.</param>
        /// <returns>The value of SUT's AccessRights property.</returns>
        [MethodHelp("Log on to the server (serverComputerName) with the specified user account (userName, userPassword, userDomain), " +
            "and get the value of SUT's AccessRights property. " +
            "Enter the current value of SUT's AccessRights property.")]
        string GetMailboxFolderPermission(string serverComputerName, string userName, string userPassword, string userDomain);

        /// <summary>
        /// Sets SUT's AccessRights property to a specified value.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the server.</param>
        /// <param name="userName">The name of the user, who is in administrators group of server, used to communicate with the server.</param>
        /// <param name="userPassword">The password of the user used to communicate with server.</param>
        /// <param name="userDomain">The domain of the user used to communicate with server.</param>
        /// <param name="permission">The new value of AccessRights.</param>
        [MethodHelp("Log on to the server (serverComputerName) with the specified user account (userName, userPassword, userDomain), " +
            "and set SUT's AccessRights property to the specified value (permission). ")]
        void SetMailboxFolderPermission(string serverComputerName, string userName, string userPassword, string userDomain, string permission);
    }
}