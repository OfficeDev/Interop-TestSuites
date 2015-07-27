//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Reflection;
    using System.Runtime.InteropServices;

    /// <summary>
    /// MS-OXCRPC 2.2.2.1 RPC_HEADER_EXT
    /// </summary>
    [StructLayout(LayoutKind.Explicit, Size = 8)]
    public struct RPC_HEADER_EXT
    {
        /// <summary>
        /// Defines the version of the header. There is only one version of the header at this time so this value MUST be set to 0x0000.
        /// </summary>
        [FieldOffset(0)]
        public ushort Version;

        /// <summary>
        /// Flags that specify how data follows this header MUST be interpreted.
        /// Compressed 0x0001  The data that follows the RPC_HEADER_EXT is compressed. The size of the data when uncompressed is in field SizeActual. If this flag is not set, the Size and SizeActual fields MUST be the same.
        /// XorMagic   0x0002  The data following the RPC_HEADER_EXT has been obfuscated. See section 3.1.7.3 for more information about the obfuscation algorithm.
        /// Last       0x0004  Indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT. This flag is used to indicate that there are multiple buffers, each with its own RPC_HEADER_EXT, one after the other.
        /// </summary>
        [FieldOffset(2)]
        public ushort Flags;

        /// <summary>
        ///  The total length of the payload data that follows the RPC_HEADER_EXT structure. This length does not include the length of the RPC_HEADER_EXT structure.
        /// </summary>
        [FieldOffset(4)]
        public ushort Size;

        /// <summary>
        /// The length of the payload data after it has been uncompressed. This field is only useful if the Compressed flag is set in the Flags field. If the Compressed flag is not set, this value MUST be equal to Size.
        /// </summary>
        [FieldOffset(6)]
        public ushort SizeActual;
    }

    /// <summary>
    /// MS-OXCRPC 2.2.2.2 AUX_HEADER
    /// </summary>
    [StructLayout(LayoutKind.Explicit, Size = 4)]
    public struct AUX_HEADER
    {
        /// <summary>
        /// Size of the AUX_HEADER structure plus any additional payload data that follows.
        /// </summary>
        [FieldOffset(0)]
        public ushort Size;

        /// <summary>
        /// Version information of the payload data that follows the AUX_HEADER. This value in conjunction with the Type field determines which structure to use to interpret the data that follows the header.
        /// AUX_VERSION_1 0x01
        /// AUX_VERSION_2 0x02
        /// </summary>
        [FieldOffset(2)]
        public byte Version;

        /// <summary>
        /// Type of payload data that follows the AUX_HEADER. This value in conjunction with the Version field determines which structure to use to interpret the data that follows the header.
        /// </summary>
        [FieldOffset(3)]
        public byte Type;
    }

    /// <summary>
    /// Section 3.1.7.1 ExtendedBuffer
    /// </summary>
    public struct ExtendedBuffer
    {
        /// <summary>
        /// The header of ExtendedBuffer
        /// </summary>
        public RPC_HEADER_EXT Header;

        /// <summary>
        /// The Payload of ExtendedBuffer
        /// </summary>
        public byte[] Payload;
    }

    /// <summary>
    ///  Class NativeMethods exposes the methods of OXCRPC
    /// </summary>
    public class NativeMethods
    {
        /// <summary>
        /// The RPC run time dll name.
        /// </summary>
        private const string RPCRuntimeDllName = "rpcrt4.dll";

        /// <summary>
        /// the file name of MS-OXCRPC Stub dll.
        /// </summary>
        private const string FileName = @"MS-OXCRPC_RPCStub.dll";

        #region Unmanaged RPC Calls
        /// <summary>
        /// The EcDoConnectEx method establishes a new Session Context with the server. The Session Context is persisted on the server until the client disconnects by using EcDoDisconnect. This method returns a Session Context Handle (CXH) to be used by a client in subsequent calls.
        /// </summary>
        /// <param name="bindingHandle">A valid RPC binding handle.</param>
        /// <param name="pcxh">On success, the server MUST return a unique value to be used as a CXH. This unique value serves as the CXH for the client. On failure, the server MUST return a zero value as the CXH.</param>
        /// <param name="userDN">User's distinguished name (DN). String containing the DN of the user who is making the EcDoConnectEx call in a directory service. Value: "/o=Microsoft/ou=First Administrative Group/cn=Recipients/cn=janedow".</param>
        /// <param name="flags">For ordinary client calls this value MUST be 0x00000000.</param>
        /// <param name="connectionModulus">The connection modulus is a client derived 32-bit hash value of the DN passed in field szUserDN and can be used by the server to decide which public folder replica to use when accessing public folder information when more than one replica of a folder exists. The hash can be used to distribute client access across replicas in a deterministic way for load balancing.</param>
        /// <param name="limit">This field is reserved. A client MUST pass a value of 0x00000000.</param>
        /// <param name="codePageId">The code page in which text data SHOULD be sent if Unicode format is not requested by the client on subsequent calls using this Session Context.</param>
        /// <param name="localIdString">The local ID for everything other than sorting.</param>
        /// <param name="localIdSort">The local ID for sorting.</param>
        /// <param name="sessionContextLink">This value is used to link the Session Context created by this call with an existing Session Context on the server. If no session linking is requested, this value will be 0xFFFFFFFF.</param>
        /// <param name="isCanConvertCodePages">The client MUST pass a value of 0x01.</param>
        /// <param name="pcmsPollsMax">The server returns the number of milliseconds that a client SHOULD wait between polling the server for event information.</param>
        /// <param name="retryTimes">The server returns the number of times a client SHOULD retry future RPC calls using the CXH returned in this call. This is for client RPC calls that fail with RPC status code RPC_S_SERVER_TOO_BUSY. This is a suggested retry count for the client and SHOULD NOT be enforced by the server.</param>
        /// <param name="pcmsRetryDelay">The server returns the number of milliseconds a client SHOULD wait before retrying a failed RPC call. If any future RPC call to the server using the CXH returned in this call fails with RPC status code RPC_S_SERVER_TOO_BUSY, it SHOULD wait the number of milliseconds specified in this output parameter before retrying the call. The number of times a client SHOULD retry is returned in parameter pcRetry. This is a suggested delay for the client and SHOULD NOT be enforced by the server.</param>
        /// <param name="picxr">The server returns a session index value that is associated with the CXH returned from this call. This value in conjunction with the session creation time stamp value returned in pulTimeStamp will be passed to a subsequent EcDoConnectEx call, if the client wants to link two Session Contexts. The server MUST NOT assign two active Session Contexts the same session index value. The server is free to return any 16-bit value for the session index.</param>
        /// <param name="valueOfDNPrefix">The server returns the distinguished name (DN) of the server.</param>
        /// <param name="displayName">The server returns the display name of the server.</param>
        /// <param name="rgwClientVersion">The client passes the client protocol version the server SHOULD use to determine what protocol functionality the client supports. For more information about how version numbers are interpreted from the wire data, see section 3.1.9.[MS-OXCRPC]</param>
        /// <param name="rgwServerVersion">The server returns the server protocol version the client SHOULD use to determine what protocol functionality the server supports. For details about how version numbers are interpreted from the wire data, see section 3.1.9.[MS-OXCRPC]</param>
        /// <param name="rgwBestVersion">The server returns the minimum client protocol version the server supports. This information is useful if the EcDoConnectEx call fails with return code ecVersionMismatch. On success, the server SHOULD return the value passed in rgwClientVersion by the client. For details about how version numbers are interpreted from the wire data, see section 3.1.9.[MS-OXCRPC]</param>
        /// <param name="pulTimeStamp">On input, this parameter and parameter ulIcxrLink are used for linking the Session Context created by this call with an existing Session Context. If the ulIcxrLink parameter is not 0xFFFFFFFF, the client MUST pass in the pulTimeStamp value returned from the server on a previous call to EcDoConnectEx (see the ulIcxrLink and piCxr parameters for more details).</param>
        /// <param name="rgbAuxIn">This parameter contains an auxiliary payload buffer. The auxiliary payload buffer is prefixed by an RPC_HEADER_EXT structure. Information stored in this header determines how to interpret the data following the header. The length of the auxiliary payload buffer that includes the RPC_HEADER_EXT header is contained in parameter cbAuxIn.</param>
        /// <param name="inputAuxSize">On input, this parameter contains the length of the auxiliary payload buffer passed in the rgbAuxIn parameter. The server MUST fail with error code ecRpcFormat if the request buffer is larger than 0x00001008 bytes in size.</param>
        /// <param name="rgbAuxOut">On output, the server can return auxiliary payload data to the client. The server MUST include an RPC_HEADER_EXT header before the auxiliary payload data.</param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer. The server MUST fail with error code ecRpcFormat if this value is larger than 0x00001008. On output, this parameter contains the size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code or one of the protocol-defined error codes listed in the following table.[MS-OXCRPC].p41. No exceptions are thrown beyond those thrown by the underlying RPC protocol [MS-RPCE].</returns>
        /// <remarks>MS-OXCRPC 3.1.4.11 EcDoConnectEx (opnum 10).</remarks>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern uint EcDoConnectEx(
                        IntPtr bindingHandle,
                        ref IntPtr pcxh,
                        string userDN,
                        uint flags,
                        uint connectionModulus,
                        uint limit,
                        uint codePageId,
                        uint localIdString,
                        uint localIdSort,
                        uint sessionContextLink,
                        ushort isCanConvertCodePages,
                        out uint pcmsPollsMax,
                        out uint retryTimes,
                        out uint pcmsRetryDelay,
                        out ushort picxr,
                        out UIntPtr valueOfDNPrefix,
                        out UIntPtr displayName,
                        ushort[] rgwClientVersion,
                        ushort[] rgwServerVersion,
                        ushort[] rgwBestVersion,
                        ref uint pulTimeStamp,
                        byte[] rgbAuxIn,
                        uint inputAuxSize,
                        byte[] rgbAuxOut,
                        ref uint pcbAuxOut);

        /// <summary>
        /// The method EcDoDisconnect closes a Session Context with the server
        /// </summary>
        /// <param name="pcxh">On input, contains the CXH of the Session Context that the client wants to disconnect. 
        /// On output, the server MUST clear the CXH to a zero value.</param>
        /// <returns>If the method succeeds, the return value is 0. 
        /// If the method fails, the return value is an implementation-specific error code</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        public static extern uint EcDoDisconnect(ref IntPtr pcxh);

        /// <summary>
        /// The method EcDoRpcExt2 passes generic remote operation (ROP) commands to the server for processing within a Session Context.
        /// </summary>
        /// <param name="pcxh">On input, the client MUST pass a valid Session Context Handle (CXH)that was created by calling EcDoConnectEx.
        /// On output, the server MUST return the same CXH on success</param>
        /// <param name="pulFlags">On input, this parameter contains flags that tell the server how to build the rgbOut parameter.</param>
        /// <param name="rgbIn">This buffer contains the ROP request payload.</param>
        /// <param name="inputRopSize">On input, this parameter contains the length of the ROP request payload passed in the rgbIn parameter.</param>
        /// <param name="rgbOut">On success, this buffer contains the ROP response payload</param>
        /// <param name="pcbOut">On input, this parameter contains the maximum size of the rgbOut buffer</param>
        /// <param name="rgbAuxIn">This parameter contains an auxiliary payload buffer.</param>
        /// <param name="inputAuxSize">On input, this parameter contains the length of the auxiliary payload buffer passed in the rgbAuxIn parameter.</param>
        /// <param name="rgbAuxOut">On output, the server can return auxiliary payload data to the client.</param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer</param>
        /// <param name="pulTransTime">On output, the server stores the number of milliseconds the call took to execute</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code or the protocol-defined error code listed in the following table.[MS-OXCRPC].p43</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern uint EcDoRpcExt2(
                        ref IntPtr pcxh,
                        ref uint pulFlags,
                        IntPtr rgbIn,
                        uint inputRopSize,
                        IntPtr rgbOut,
                        ref uint pcbOut,
                        IntPtr rgbAuxIn,
                        uint inputAuxSize,
                        IntPtr rgbAuxOut,
                        ref uint pcbAuxOut,
                        out uint pulTransTime);

        /// <summary>
        /// A client can use it to determine if it can communicate with the server.
        /// </summary>
        /// <param name="bindingHandle">A valid RPC binding handle</param>
        /// <returns>The function MUST always succeed and return 0.</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        public static extern uint EcDummyRpc(IntPtr bindingHandle);

        /// <summary>
        /// The method EcDoAsyncConnectEx binds a Session Context Handle (CXH) returned from method 
        /// EcDoConnectEx to a new Asynchronous Context Handle (ACXH) that can be used in calls to EcDoAsyncWaitEx in interface AsyncEMSMDB.
        /// </summary>
        /// <param name="pcxh">Client MUST pass a valid CXH that was created by calling EcDoConnectEx</param>
        /// <param name="pacxh">On success, the server returns an ACXH that is associated with the Session Context passed in parameter CXH. 
        /// On a failure the returned value is a Null.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code or the protocol-defined error code listed in the following table. [MS-OXCRPC].p43</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern uint EcDoAsyncConnectEx(
                        IntPtr pcxh,
                        ref IntPtr pacxh);

        /// <summary>
        /// The method EcDoAsyncWaitEx is an asynchronous call that the server will not complete until there are pending events on the Session Context up to the duration which the input parameter waitSecondThreshold specified. 
        /// </summary>
        /// <param name="pacxh">a unique value to be used as a ACXH</param>
        /// <param name="inputFlags">Unused.Reserved for future use.Client MUST pass a value of 0x00000000.</param>
        /// <param name="waitSecondThreshold">Indicates the threshold of waiting time in second</param>
        /// <param name="makeEvent">it indicates whether Client sends a event to Server</param>
        /// <param name="pulFlagsOut">Output flags for the client.</param>
        /// <returns>If success, it returns 0, else returns the error code</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern uint EcDoAsyncWaitExWrap(
            IntPtr pacxh,
            uint inputFlags,
            uint waitSecondThreshold,
            [MarshalAs(UnmanagedType.Bool)]
            bool makeEvent,
            out uint pulFlagsOut);

        /// <summary>
        /// The method EcRRegisterPushNotification registers a callback address with the server for a Session Context.
        /// </summary>
        /// <param name="pcxh">On input, the client MUST pass a valid CXH that was created by calling EcDoConnectEx</param>
        /// <param name="valueOfiRpc">The server MUST completely ignore this value.</param>
        /// <param name="rgbContext">This parameter contains opaque client-generated context data that is sent back to the client at the callback address</param>
        /// <param name="clientContextSize">This parameter contains the size of the opaque client context data that is passed in parameter rgbContext.</param>
        /// <param name="grbitAdviseBits">This parameter MUST be 0xFFFFFFFF</param>
        /// <param name="rgbCallbackAddress">This parameter contains the callback address for the server to use to notify the client of a pending event</param>
        /// <param name="callbackAddressSize">This parameter contains the length of the callback address in parameter rgbCallbackAddress.</param>
        /// <param name="notificationHandle">If the call completes successfully, this output parameter will contain a handle to the notification callback on the server</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code or one of the protocol-defined error codes listed in the following table. MS-OXCRPC.p35</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern uint EcRRegisterPushNotification(
                        ref IntPtr pcxh,
                        ulong valueOfiRpc,
                        IntPtr rgbContext,
                        ushort clientContextSize,
                        ulong grbitAdviseBits,
                        IntPtr rgbCallbackAddress,
                        ushort callbackAddressSize,
                        ref ulong notificationHandle);

        /// <summary>
        /// The method EcRRegisterPushNotificationWrap registers a callback address with the server for a Session Context.
        /// </summary>
        /// <param name="pcxh">On input, the client MUST pass a valid CXH that was created by calling EcDoConnectEx</param>
        /// <param name="family">Ip address family</param>
        /// <param name="valueOfIPAddr">Ip address</param>
        /// <param name="port">Port number of a socket</param>
        /// <param name="rgbContext">This parameter contains opaque client-generated context data that is sent back to the client at the callback address</param>
        /// <param name="clientContextSize">This parameter contains the size of the opaque client context data that is passed in parameter rgbContext.</param>
        /// <param name="notificationHandle">If the call completes successfully, this output parameter will contain a handle to the notification callback on the server</param>
        /// <returns>If success, it returns 0, else returns the error code</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern uint EcRRegisterPushNotificationWrap(
            ref IntPtr pcxh,
            short family,
            string valueOfIPAddr,
            ushort port,
            byte[] rgbContext,
            ushort clientContextSize,
            out uint notificationHandle);

        /// <summary>
        /// This method binds client to RPC server.
        /// </summary>
        /// <param name="serverName">Representation of a network address of server.</param>
        /// <param name="encryptionMethod">The encryption method in this call.</param>
        /// <param name="authnSvc">Authentication service to use.</param>
        /// <param name="seqType">Transport sequence type</param>
        /// <param name="rpchUseSsl">true to use RPC over HTTP with SSL, false to use RPC over HTTP without SSL.</param>
        /// <param name="rpchAuthScheme">The authentication scheme used in the http authentication for RPC over HTTP. This value can be "Basic" or "NTLM".</param>
        /// <param name="spnStr">Service Principal Name (SPN) string used in Kerberos SSP</param>
        /// <param name="options">proxy attribute</param>
        /// <param name="setUuid">True to set PFC_OBJECT_UUID(0x80) field of RPC header, false to not set this field</param>
        /// <returns>Binding status.The non-zero return value indicates failed binding.</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern uint BindToServer(string serverName, uint encryptionMethod, uint authnSvc, string seqType, [MarshalAs(UnmanagedType.Bool)]bool rpchUseSsl, string rpchAuthScheme, string spnStr, string options, [MarshalAs(UnmanagedType.Bool)]bool setUuid);

        /// <summary>
        /// Create SEC_WINNT_AUTH_IDENTITY structure in native codes that enables passing a 
        /// particular user name and password to the run-time library for the purpose of authentication.
        /// </summary>
        /// <param name="domain">String containing the domain or workgroup name.</param>
        /// <param name="userName">String containing the user name.</param>
        /// <param name="password">String containing the user's password in the domain or workgroup.</param>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern void CreateIdentity(string domain, string userName, string password);

        /// <summary>
        /// Return the current binding handle.
        /// </summary>
        /// <returns>Current binding handle.</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr GetBindHandle();

        /// <summary>
        /// Complete the rpc async call
        /// </summary>
        /// <param name="pAsync">The rpc async handle</param>
        /// <param name="pReply">The rpc call return value</param>
        /// <returns>The status of asynchronous call</returns>
        [DllImport(RPCRuntimeDllName)]
        public static extern int RpcAsyncCompleteCall(
            IntPtr pAsync,
            IntPtr pReply);

        /// <summary>
        /// Query the rpc async call status
        /// </summary>
        /// <param name="pAsync">The rpc async handle</param>
        /// <returns>The status of asynchronous call</returns>
        [DllImport(RPCRuntimeDllName)]
        public static extern int RpcAsyncGetCallStatus(IntPtr pAsync);

        /// <summary>
        /// Create a RPC handle for async use
        /// </summary>
        /// <returns>Return the pointer</returns>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr CreateRpcAsyncHandle();

        /// <summary>
        /// Call EcDoAsyncWaitEx method and return immediately.
        /// </summary>
        /// <param name="ecDoAsyncWaitEx_AsyncHandle">RPC async handle</param>
        /// <param name="acxh">The asynchronous context handle</param>
        /// <param name="r">Unused. Reserved for future use. Client MUST pass a value of 0x00000000</param>
        /// <param name="o">Output flags for the client</param>
        [DllImport(FileName, CallingConvention = CallingConvention.StdCall)]
        public static extern void EcDoAsyncWaitEx(
            IntPtr ecDoAsyncWaitEx_AsyncHandle,
            IntPtr acxh,
            int r,
            IntPtr o);

        #endregion

        /// <summary>
        /// Get the RPC Exception Code
        /// </summary>
        /// <param name="e">
        /// Exception object</param>
        /// <returns>Returns the RPC error code</returns>
        public static uint RpcExceptionCode(SEHException e)
        {
            uint errorCode = 0;
            Type sehType = typeof(SEHException);
            FieldInfo xcodeField = sehType.BaseType.BaseType.BaseType.GetField("_xcode", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (xcodeField != null)
            {
                int code = (int)xcodeField.GetValue(e);
                errorCode = (uint)code;
            }

            return errorCode;
        }

        #region RPC Runtime Methods

        /// <summary>
        /// Free handle
        /// </summary>
        /// <param name="binding"> A pointer to the server binding handle</param>
        /// <returns>Returns an integer value to indicate call success or failure</returns>
        [DllImport(RPCRuntimeDllName)]
        public static extern uint RpcBindingFree(
            ref IntPtr binding);

        /// <summary>
        /// Enables client applications to specify message-queuing options on a binding handle.
        /// </summary>
        /// <param name="binding">Server binding to modify.</param>
        /// <param name="option">Binding property to modify.</param>
        /// <param name="optionValue">New value for the binding property. </param>
        /// <returns>Returns an integer value to indicate call success or failure</returns>
        [DllImport(RPCRuntimeDllName)]
        public static extern uint RpcBindingSetOption(
            IntPtr binding,
            uint option,
            uint optionValue);
        #endregion
    }
}