//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXCNOTIF Protocol Adapter Interface.
    /// </summary>
    public interface IMS_OXCNOTIFAdapter : IAdapter
    {
        /// <summary>
        /// Gets or sets the value of the RPC context returned from the last connect.
        /// </summary>
        IntPtr RPCContext { get; set; }

        /// <summary>
        /// Gets or sets the value of the LogonHandle returned from the last connect.
        /// </summary>
        uint LogonHandle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the session is connect.
        /// </summary>
        bool IsConnected { get; set; }

        /// <summary>
        /// Logon to server by calling RopLogon.
        /// </summary>
        /// <returns>The server response.</returns>
        RopLogonResponse Logon();

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection.</param>
        /// <returns>If the behavior of connecting server is successful, the server will return true; otherwise, return false.</returns>
        bool DoConnect(ConnectionType connectionType);

        /// <summary>
        /// Disconnect the connection with Server.
        /// </summary>
        /// <returns>The return value indicates the disconnection state.</returns>
        bool DoDisconnect();

        /// <summary>
        /// Switch the instance of OxcropsClient.
        /// </summary>
        void SwitchSessionContext();

        /// <summary>
        /// Register a callback address on the server which will be used to push notification.
        /// </summary>
        /// <param name="addressFamily">The addressFamily specifies which IP family to use.</param>
        /// <param name="port">The UDP port that will receive the push notification.</param>
        /// <param name="opaque">The opaque client-generated context data that is sent back to the client at the callback address.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code.</returns>
        uint EcRRegisterPushNotification(AddressFamily addressFamily, int port, string opaque);

        /// <summary>
        /// Receive push notification on the specified port.
        /// </summary>
        /// <param name="addressFamily">The addressFamily specifies which IP family to use.</param>
        /// <param name="port">The UDP port uses to receive the push notification.</param>
        /// <param name="opaque">The opaque data received from server.</param>
        /// <returns>True if the notification received, otherwise false.</returns>
        bool PushNotificationReceived(AddressFamily addressFamily, int port, out string opaque);

        /// <summary>
        /// Acquire an asynchronous context handle on the server which will be used in subsequent EcDoAsyncWaitEx call.
        /// </summary>
        /// <returns>The asynchronous context handle.</returns>
        IntPtr EcDoAsyncConnectEx();

        /// <summary>
        /// Call EcDoAsyncWaitEx method and return immediately.
        /// </summary>
        /// <param name="acxh">The asynchronous context handle</param>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        void BeginAsyncWait(IntPtr acxh, out IntPtr rpcAsyncHandle);

        /// <summary>
        /// Get the status of EcDoAsyncWaitEx call.
        /// </summary>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        /// <returns>The status of asynchronous call</returns>
        RPCAsyncStatus QueryAsyncWaitStatus(IntPtr rpcAsyncHandle);

        /// <summary>
        /// Complete the EcDoAsyncWaitEx call.
        /// </summary>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        /// <param name="pulFlagsOut">The pulFlagsOut parameter returned by EcDoAsyncWaitEx</param>
        /// <returns>The EcDoAsyncWaitEx call return value</returns>
        int EndAsyncWait(IntPtr rpcAsyncHandle, out int pulFlagsOut);

        /// <summary>
        /// Create a subscription for the specified notifications on the server. 
        /// </summary>
        /// <param name="notificationType">The notification type which want to subscribe.</param>
        /// <returns>The server response.</returns>
        RopRegisterNotificationResponse RegisterNotification(NotificationType notificationType);

        /// <summary>
        /// Create a subscription for the specified notifications on the server with specified folder ID and message ID. 
        /// </summary>
        /// <param name="notificationType">The notification type which want to subscribe.</param>
        /// <param name="wantWholeStore">The value of WantWholeStore.</param>
        /// <param name="flolderId">The value of specified folder ID.</param>
        /// <param name="messageId">The value of specified message ID</param>
        /// <param name="notificationHandle">The value of notificationHandle</param>
        /// <returns>The server response.</returns>
        RopRegisterNotificationResponse RegisterNotificationWithParameter(NotificationType notificationType, byte wantWholeStore, ulong flolderId, ulong messageId, out uint notificationHandle);

        /// <summary>
        /// Retrieve the RopNotify and RopPending response by sending an empty RPC request.
        /// </summary>
        /// <param name="expectGetNotification">Whether expect to get any notification.</param>
        /// <returns>The list of ROP response, empty if no response returned.</returns>
        IList<IDeserializable> GetNotification(bool expectGetNotification);

        /// <summary>
        /// Send ROP request with single operation and single input object handle.
        /// </summary>
        /// <param name="requestRop">The ROP request.</param>
        /// <param name="inputObjHandles">The input object handle.</param>
        /// <param name="responseSOHs">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <returns>The responses returned by the server.</returns>
        IList<IDeserializable> Process(ISerializable requestRop, uint inputObjHandles, out List<List<uint>> responseSOHs);

        /// <summary>
        /// Send ROP request with single operation, multiple input object handles.
        /// </summary>
        /// <param name="ropRequest">The ROP request.</param>
        /// <param name="inputObjHandles">The multiple input object handles.</param>
        /// <param name="responseSOHs">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <returns>The responses returned by the server.</returns>
        IList<IDeserializable> Process(ISerializable ropRequest, IEnumerable<uint> inputObjHandles, out List<List<uint>> responseSOHs);

        /// <summary>
        /// The method to send NotificationWait request to the server.
        /// </summary>        
        /// <param name="requestBody">The NotificationWait request body.</param>
        /// <returns>Return the NotificationWait response from the server.</returns>
        NotificationWaitSuccessResponseBody NotificationWait(NotificationWaitRequestBody requestBody);
    }
}