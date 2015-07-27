//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-ASPROV SUT control adapter interface.
    /// </summary>
    public interface IMS_ASPROVSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Wipe all data of the user's device.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the SUT.</param>
        /// <param name="userEmail">The email address of the user whose device will be remote wiped.</param>
        /// <param name="userPassword">The password of the user whose device will be remote wiped.</param>
        /// <param name="deviceType">The DeviceType of the device which will be remote wiped.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to the SUT (serverComputerName) with the specified user account (userEmail, userPassword), " +
            "and request a remote wipe operation to wipe all data on the specified device (deviceType). " +
            "If the operation succeeds, enter \"true\"; otherwise, enter \"false\".")]
        bool WipeData(string serverComputerName, string userEmail, string userPassword, string deviceType);

        /// <summary>
        /// Remove the device from the user's mobile list.
        /// </summary>
        /// <param name="serverComputerName">The computer name of the SUT.</param>
        /// <param name="userEmail">The email address of the user whose device will be removed.</param>
        /// <param name="userPassword">The password of the user whose device will be removed.</param>
        /// <param name="deviceType">The DeviceType of the device which will be removed.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to the SUT (serverComputerName) with the specified user account (userEmail, userPassword), " +
            "and remove the specified device (deviceType) from the mobile list. " +
            "If the operation succeeds, enter \"true\"; otherwise, enter \"false\".")]
        bool RemoveDevice(string serverComputerName, string userEmail, string userPassword, string deviceType);
    }
}