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
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of the MS-OXCSTOR Adapter
    /// </summary>
    public interface IMS_OXCSTORAdapter : IAdapter
    {
        /// <summary>
        /// Connect to server for RPC calling. 
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <returns>True indicates connecting successfully, otherwise false</returns>
        bool ConnectEx(ConnectionType connectionType);

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">The type of connection</param>
        /// <param name="userDN">UserDN used to connect server</param>
        /// <param name="domain">The domain the server is deployed</param>
        /// <param name="userName">The domain account name</param>
        /// <param name="password">User password</param>
        /// <returns>True indicates connecting successfully, otherwise false</returns>
        bool ConnectEx(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password);

        /// <summary>
        /// Disconnect the connection with server.
        /// </summary>
        /// <returns>True indicates disconnecting successfully, otherwise false</returns>
        bool DisconnectEx();

        /// <summary>
        /// Send ROP request with single operation with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="commandType">ROP commands type</param>
        /// <param name="outputBuffer">ROP response buffer</param>
        void DoRopCall(ISerializable ropRequest, uint inputObjHandle, ROPCommandType commandType, out RopOutputBuffer outputBuffer);

        /// <summary>
        /// Set auto redirect value in RPC context
        /// </summary>
        /// <param name="option">True indicates enable auto redirect, false indicates disable auto redirect</param>
        void SetAutoRedirect(bool option);
    }
}