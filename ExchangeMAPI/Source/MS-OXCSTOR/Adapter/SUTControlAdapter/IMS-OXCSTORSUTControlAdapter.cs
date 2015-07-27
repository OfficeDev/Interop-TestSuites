//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// SUT control adapter which carries out various operations to configure the SUT.
    /// </summary>
    public interface IMS_OXCSTORSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Enable a mailbox on the SUT.
        /// </summary>
        /// <param name="userName">A string identifies the mailbox object to enable.</param>
        /// <returns>If the operation succeeded, return "success"; otherwise return the error information.</returns>
        [MethodHelp("Enable the mailbox specified in the parameter userName in  \"SutComputerName\". The SutComputerName property specifies the SUT name in the ExchangeCommonConfiguration.deployment.ptfconfig file. " + "\r\n" +
                    "userName: The user whose mailbox needs to be enabled. " + "\r\n" +
                    "Return value: If the operation has succeeded, return \"success\"; otherwise return the error information.")]
        string EnableMailbox(string userName);

        /// <summary>
        /// Disable a mailbox on the SUT.
        /// </summary>
        /// <param name="userName">A string identifies the mailbox object to disable.</param>
        /// <returns>If the operation succeeded, return "success"; otherwise return the error information.</returns>
        [MethodHelp(" Disable the mailbox specified in the parameter userName in  \"SutComputerName\". The SutComputerName property specifies the SUT name in the ExchangeCommonConfiguration.deployment.ptfconfig file. " + "\r\n" +
                    "userName: The user whose mailbox needs to be disabled. " + "\r\n" +
                    "Return value: If the operation has succeeded, return \"success\"; otherwise return the error information.")]
        string DisableMailbox(string userName);

        /// <summary>
        /// Get the value of LegacyExchangeDN of a mailbox user.
        /// </summary>
        /// <param name="computerName">The computer name of the server.</param>
        /// <param name="userName">The user whose LegacyExchangeDN is returned.</param>
        /// <returns>If the operation succeeded, return the value of LegacyExchangeDN; otherwise return null or empty.</returns>
        [MethodHelp("Get the value of the LegacyExchangeDN of a mailbox user." + "\r\n" +
                    "computerName: The computer name of the server." + "\r\n" +
                    "userName: The user whose LegacyExchangeDN value is returned." + "\r\n" +
                    "Return value: If the LegacyExchangeDN value exists, enter the value; otherwise just keep the Return Value field empty." + "\r\n" +
                    "If the server is a Microsoft Exchange server, the value of LegacyExchangeDN can be obtained from Active Directory Service Interfaces Editor (ADSI Edit).")]
        string GetUserDN(string computerName, string userName);
    }
}