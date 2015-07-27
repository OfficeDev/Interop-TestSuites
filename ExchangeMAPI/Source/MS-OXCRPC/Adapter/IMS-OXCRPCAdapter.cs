//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// An interface defines the protocol adapter methods used by the MS-OXCRPC test cases.
    /// </summary>
    public interface IMS_OXCRPCAdapter : IAdapter
    {
        /// <summary>
        /// Initializes the client and server, build the transport tunnel between client and server.
        /// </summary>
        /// <param name="encryptionMethod">An unsigned integer indicates the authentication level for creating RPC binding</param>
        /// <param name="authnSvc">An unsigned integer indicates authentication services.</param>
        /// <param name="userName">Define user name which can be used by client to access SUT. </param>
        /// <param name="password">Define user password which can be used by client to access SUT.</param>
        /// <returns>If success, it returns true, else returns false.</returns>
        bool InitializeRPC(uint encryptionMethod, uint authnSvc, string userName, string password);

        /// <summary>
        /// The method EcRRegisterPushNotification registers a callback address with the server for a Session Context. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="rgbContext">This parameter contains opaque client-generated context data that is sent back to the client at the callback address.</param>
        /// <param name="addType">The type of the cbCallbackAddress.</param>
        /// <param name="ip">The client IP used in this method.</param>
        /// <param name="notificationHandle">If the call completes successfully, this output parameter will contain a handle to the notification callback on the server.</param>
        /// <returns>If success, it returns 0, else returns the error code.</returns>
        uint EcRRegisterPushNotification(
            ref IntPtr pcxh, 
            byte[] rgbContext, 
            Add_Families addType, 
            string ip, 
            out uint notificationHandle);

        /// <summary>
        /// The method EcDoRpcExt2 passes generic remote operation (ROP) commands to the server for processing within a Session Context. Each call can contain multiple ROP commands. 
        /// </summary>
        /// <param name="pcxh">The unique value points to a CXH.</param>
        /// <param name="pulFlags">Flags that tell the server how to build the rgbOut parameter.</param>
        /// <param name="rgbIn">This buffer contains the ROP request payload. </param>
        /// <param name="pcbOut">On input, this parameter contains the maximum size of the rgbOut buffer. On output, this parameter contains the size of the ROP response payload.</param>
        /// <param name="rgbAuxIn">This parameter contains an auxiliary payload buffer. </param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer. On output, this parameter contains the size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="response">The ROP response returned in this RPC method and parsed by adapter.</param>
        /// <param name="responseSOHTable">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoRpcExt2(
            ref IntPtr pcxh, 
            PulFlags pulFlags, 
            byte[] rgbIn, 
            ref uint pcbOut, 
            byte[] rgbAuxIn, 
            ref uint pcbAuxOut,
            out IDeserializable response, 
            ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// The method EcDoRpcExt2 passes generic remote operation (ROP) commands to the server for processing within a Session Context. Each call can contain multiple ROP commands. 
        /// </summary>
        /// <param name="pcxh">The unique value to be used as a CXH</param>
        /// <param name="pulFlags">Flags that tell the server how to build the rgbOut parameter.</param>
        /// <param name="rgbIn">This buffer contains the ROP request payload. </param>
        /// <param name="rgbOut">On Output, this parameter contains the response payload.</param>
        /// <param name="pcbOut">On Output, this parameter contains the size of response payload.</param>
        /// <param name="rgbAuxIn">This parameter contains an auxiliary payload buffer. </param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer. On output, this parameter contains the size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="response">The ROP response returned in this RPC method and parsed by adapter.</param>
        /// <param name="responseSOHTable">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <param name="payloadCount">The count of payload that ROP response contains.</param>
        /// <param name="rgbAuxOut">On output, this parameter contains auxiliary payload data.</param>
        /// <returns>If success, it return 0, else return the error code.</returns>
        uint EcDoRpcExt2(
            ref IntPtr pcxh, 
            PulFlags pulFlags, 
            byte[] rgbIn, 
            ref byte[] rgbOut, 
            ref uint pcbOut, 
            byte[] rgbAuxIn, 
            ref uint pcbAuxOut,
            out IDeserializable response, 
            ref List<List<uint>> responseSOHTable, 
            out uint payloadCount,
            ref byte[] rgbAuxOut);

        /// <summary>
        /// The method EcDoDisconnect closes the Session Context with the server. 
        /// </summary>
        /// <param name="pcxh">The unique value points to a CXH.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoDisconnect(ref IntPtr pcxh);

        /// <summary>
        /// The method EcDoAsyncConnectEx binds a Session Context Handle (CXH) returned from method EcDoConnectEx to a new Asynchronous Context Handle (ACXH) 
        /// that can be used in calls to EcDoAsyncWaitEx in interface AsyncEMSMDB. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="pacxh">An ACXH that is associated with the Session Context passed in parameter CXH.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoAsyncConnectEx(IntPtr pcxh, ref IntPtr pacxh);

        /// <summary>
        /// This RPC method determines if it can communicate with the server.
        /// </summary>
        /// <returns>Always return zero.</returns>
        uint EcDummyRpc();

        /// <summary>
        /// The method EcDoAsyncWaitEx is an asynchronous call that will not be complete by server until there are pending events on the Session Context up to a five minute duration. 
        /// </summary>
        /// <param name="acxh">The unique value to be used as a CXH.</param>
        /// <param name="isNotificationPending">A Boolean value indicates signals that events are pending for the client on the Session Context on the server. </param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoAsyncWaitEx(IntPtr acxh, out bool isNotificationPending);

        /// <summary>
        /// The EcDoConnectEx method establishes a new Session Context with the server. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="sessionContextLink">This value is used to link the Session Context created by this call with an existing Session Context on the server.</param>
        /// <param name="pulTimeStamp">The server has to return a time stamp in which the new Session Context was created.</param>
        /// <param name="rgbAuxIn">The auxiliary payload data.</param>
        /// <param name="userDN">The userDN input by client.</param>
        /// <param name="pcbAuxOut">The maximum length of the rgbAuxOut buffer.</param>
        /// <param name="rgwClientVersion">The client protocol version.</param>
        /// <param name="rgwBestVersion">The minimum client protocol version that the server supports.</param>
        /// <param name="picxr">The session index value that is associated with the CXH returned from this call.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoConnectEx(
            ref IntPtr pcxh, 
            uint sessionContextLink, 
            ref uint pulTimeStamp, 
            byte[] rgbAuxIn, 
            string userDN, 
            ref uint pcbAuxOut, 
            ushort[] rgwClientVersion, 
            out ushort[] rgwBestVersion, 
            out ushort picxr);

        /// <summary>
        /// The EcDoConnectEx method establishes a new Session Context with the server. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="sessionContextLink">This value is used to link the Session Context created by this call with an existing Session Context on the server.</param>
        /// <param name="pulTimeStamp">The server has to return a time stamp in which the new Session Context was created.</param>
        /// <param name="rgbAuxIn">The auxiliary payload data.</param>
        /// <param name="userDN">The userDN input by client.</param>
        /// <param name="pcbAuxOut">The maximum length of the rgbAuxOut buffer.</param>
        /// <param name="rgwClientVersion">The client protocol version.</param>
        /// <param name="rgwServerVersion">The server protocol version returned by Exchange Server.</param>
        /// <param name="rgwBestVersion">The minimum client protocol version that the server supports.</param>
        /// <param name="picxr">The session index value that is associated with the CXH returned from this call.</param>
        /// <param name="flags">The ulFlags parameter of EcDoConnectEx</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoConnectEx(
            ref IntPtr pcxh, 
            uint sessionContextLink, 
            ref uint pulTimeStamp, 
            byte[] rgbAuxIn, 
            string userDN, 
            ref uint pcbAuxOut, 
            ushort[] rgwClientVersion, 
            out ushort[] rgwServerVersion, 
            out ushort[] rgwBestVersion, 
            out ushort picxr,
            uint flags);

         /// <summary>
        /// The EcDoConnectEx method establishes a new Session Context with the server. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="sessionContextLink">This value is used to link the Session Context created by this call with an existing Session Context on the server.</param>
        /// <param name="pulTimeStamp">The server has to return a time stamp in which the new Session Context was created.</param>
        /// <param name="rgbAuxIn">The auxiliary payload data.</param>
        /// <param name="userDN">The userDN input by client.</param>
        /// <param name="pcbAuxOut">The maximum length of the rgbAuxOut buffer.</param>
        /// <param name="rgwClientVersion">The client protocol version.</param>
        /// <param name="rgwServerVersion">The server protocol version returned by Exchange Server.</param>
        /// <param name="rgwBestVersion">The minimum client protocol version that the server supports.</param>
        /// <param name="picxr">The session index value that is associated with the CXH returned from this call.</param>
        /// <param name="flags">The ulFlags parameter of EcDoConnectEx</param>
        /// <param name="rgbAuxOutValue">The additional data in the auxiliary buffers of the method.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        uint EcDoConnectEx(ref IntPtr pcxh, uint sessionContextLink, ref uint pulTimeStamp, byte[] rgbAuxIn, string userDN, ref uint pcbAuxOut, ushort[] rgwClientVersion, out ushort[] rgwServerVersion, out ushort[] rgwBestVersion, out ushort picxr, uint flags, out List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue);
    }
}