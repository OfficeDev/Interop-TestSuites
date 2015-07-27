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
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that implements RPC communications by calling native methods generated from IDL.
    /// </summary>
    public class RpcAdapter
    {
        /// <summary>
        /// This is used to set the max size of the rgbAuxOut.
        /// </summary>
        public const uint PcbAuxOut = 0x1008;

        /// <summary>
        /// This flags indicates client requests server to not compress or XOR payload of rgbOut and rgbAuxOut.
        /// </summary>
        public const uint PulFlags = 0x00000003;

        /// <summary>
        /// The binding handle.
        /// </summary>
        private IntPtr bindingHandle = IntPtr.Zero;

        /// <summary>
        /// An instance of ITestSite
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// Initializes a new instance of the RpcAdapter class
        /// </summary>
        /// <param name="site">An instance of ITestSite</param>
        public RpcAdapter(ITestSite site)
        {
            this.site = site;
        }

        /// <summary>
        /// Sends remote operation (ROP) commands to the server with a Session Context Handle.
        /// </summary>
        /// <param name="pcxh">On input, the client MUST pass a valid Session Context Handle that was created by calling EcDoConnectEx. 
        /// The server uses the Session Context Handle to identify the Session Context to use for this call. On output, the server MUST return the same Session Context Handle on success.</param>
        /// <param name="rgbIn">This buffer contains the ROP request payload. </param>
        /// <param name="rgbOut">On success, this buffer contains the ROP response payload.</param>
        /// <param name="pcbOut">On input, this parameter contains the maximum size of the rgbOut buffer.On output, this parameter contains the size of the ROP response payload.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, return RPC error code RPC exception code.</returns>
        public uint RpcExt2(
                       ref IntPtr pcxh,
                       byte[] rgbIn,
                       out byte[] rgbOut,
                       ref uint pcbOut)
        {
            uint pulFlags = PulFlags;
            uint pcbAuxOut = PcbAuxOut;
            uint pulTransTime;
            byte[] rgbAuxIn = { };
            byte[] rgbAuxOut = new byte[pcbAuxOut];
            uint ret = this.RpcExt2(
                        ref pcxh,
                        ref pulFlags,
                        rgbIn,
                        (uint)rgbIn.Length,
                        out rgbOut,
                        ref pcbOut,
                        rgbAuxIn,
                        (uint)rgbAuxIn.Length,
                        out rgbAuxOut,
                        ref pcbAuxOut,
                        out pulTransTime);
            return ret;
        }

        /// <summary>
        /// It establishes a new Session Context with the server by calling native methods.
        /// </summary>
        /// <param name="server">The server name.</param>
        /// <param name="domain">The domain the server is deployed</param>
        /// <param name="userName">The domain account name.</param>
        /// <param name="userDN">User's distinguished name (DN).</param>
        /// <param name="password">>user Password.</param>
        /// <param name="userSpn">User's SPN.</param>
        /// <param name="mapiContext">The default parameters for rpc call, such as authentication level, authentication method, whether to compress request.</param>
        /// <param name="options">proxy attribute</param>
        /// <returns>n success, the server MUST return a unique value to be used as a CXH. This unique value serves as the CXH for the client</returns>
        public IntPtr Connect(string server, string domain, string userName, string userDN, string password, string userSpn, MapiContext mapiContext, string options)
        {
            // The default parameter for out session handle.
            IntPtr pcxh = IntPtr.Zero;

            // CreateIdentity
            NativeMethods.CreateIdentity(
                domain,
                userName,
                password);

            uint status = NativeMethods.BindToServer(server, mapiContext.AuthenLevel, mapiContext.AuthenService, mapiContext.TransportSequence.ToLower(), mapiContext.RpchUseSsl, mapiContext.RpchAuthScheme, userSpn, options, mapiContext.SetUuid);
            if (status != 0)
            {
                throw new Exception("Could not create binding handle with server");
            }

            this.bindingHandle = NativeMethods.GetBindHandle();

            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.site));
            int maxRetryCount = int.Parse(Common.GetConfigurationPropertyValue("ConnectRetryCount", this.site));

            int retryCount = 0;
            do
            {
                status = this.Connect_Internal(ref pcxh, userDN, ref mapiContext);
                if (status >= 1700 && status <= 1799)
                {
                    // If status is between 1700 and 1799, try to connect again.
                    retryCount++;
                    System.Threading.Thread.Sleep(waitTime);

                    if (retryCount > 0)
                    {
                        this.site.Log.Add(LogEntryKind.Comment, "Can't connect to RPC server, will try to connect the server again. Current retry number is {0}.", retryCount);
                    }
                }
                else
                {
                    break;
                }
            }
            while (retryCount < maxRetryCount);

            if (status != 0)
            {
                string errorCodeHexString = "0x" + status.ToString("X8");
                string errorCodeMeaning = GetErrorCodeMeaning(errorCodeHexString);
                string errorCodeDescription = string.Empty;
                if (string.IsNullOrEmpty(errorCodeMeaning))
                {
                    errorCodeDescription = string.Format("Error code '{0}' is not defined in protocol MS-OXCRPC. The error message is: {1}", errorCodeHexString, (new Win32Exception((int)status)).ToString());
                }
                else
                {
                    errorCodeDescription = string.Format("Error code '{0}' is defined in protocol MS-OXCRPC as: {1}", errorCodeHexString, errorCodeMeaning);
                }

                this.site.Assert.Fail("Connect method returned an error: {0}. {1}", errorCodeHexString, errorCodeDescription);
            }

            return pcxh;
        }

        /// <summary>
        /// It calls the native method EcDoDisconnect that closes a Session Context with the server. 
        /// The Session Context is destroyed and all associated server state, objects, and resources that are associated with the Session Context are released.
        /// </summary>
        /// <param name="cxh">On input, contains the CXH of the Session Context that the client wants to disconnect. 
        ///     On output, the server MUST clear the CXH to a zero value.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code.</returns>
        public uint Disconnect(ref IntPtr cxh)
        {
            try
            {
                uint returnValue = NativeMethods.EcDoDisconnect(ref cxh);

                if (this.bindingHandle != IntPtr.Zero)
                {
                    bool rpcForceShutdownAssociation = bool.Parse(Common.GetConfigurationPropertyValue("RpcForceShutdownAssociation", this.site));
                    if (rpcForceShutdownAssociation)
                    {
                        uint status = NativeMethods.RpcBindingSetOption(this.bindingHandle, 13, 1); // 13 represents RPC_C_OPT_DONT_LINGER option
                        if (status != 0)
                        {
                            this.site.Assert.Fail("Failed to set option on the binding handle, the RpcBindingSetOption method returned status code: {0}", status);
                        }
                    }

                    NativeMethods.RpcBindingFree(ref this.bindingHandle);
                    this.bindingHandle = IntPtr.Zero;
                }

                return returnValue;
            }
            catch (SEHException e)
            {
                this.site.Log.Add(LogEntryKind.Comment, "EcDoDisconnect throws exception, system error code is {0}, the error message is: {1}", RpcExceptionCode(e), (new Win32Exception((int)RpcExceptionCode(e))).ToString());

                // The exception in ECDoDisconnect should be ignored here.
                return 0;
            }
        }

        /// <summary>
        /// Get the RPC Exception Code
        /// </summary>
        /// <param name="e">Exception object</param>
        /// <returns> 
        /// returns the RPC error code</returns>
        private static uint RpcExceptionCode(SEHException e)
        {
            uint errorCode = 0;
            Type sehType = typeof(SEHException);
            System.Reflection.FieldInfo xcodeField = sehType.BaseType.BaseType.BaseType.GetField(
                "_xcode",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (xcodeField != null)
            {
                int code = (int)xcodeField.GetValue(e);
                errorCode = (uint)code;
            }

            return errorCode;
        }

        /// <summary>
        /// Get meaning of the error code based on description in [MS-OXCRPC] section 3.1.4.11.
        /// </summary>
        /// <param name="errorCodeHexString">The error code in a hexadecimal string format. E.g. "0x80070005".</param>
        /// <returns>Returns the error code meaning as specified in [MS-OXCRPC], or returns an empty string if the error code is not defined in [MS-OXCRPC].</returns>
        private static string GetErrorCodeMeaning(string errorCodeHexString)
        {
            // Return error code meaning based on description in [MS-OXCRPC] section 3.1.4.11.
            switch (errorCodeHexString)
            {
                case "0x00000000":
                    return "Success.";
                case "0x80070005":
                    return "(ecAccessDenied) The authentication context associated with the binding handle does not have enough privilege or the szUserDN parameter is empty.";
                case "0x00000970":
                    return "(ecNotEncrypted) The server is configured to require encryption and the authentication for binding handle contained in the hBinding parameter is not set with RPC_C_AUTHN_LEVEL_PKT_PRIVACY. For more information about setting the authentication and authorization, see [MSDN-RpcBindingSetAuthInfoEx]. The client attempts the call again with new binding handle that is encrypted.";
                case "0x000004DF":
                    return "(ecClientVerDisallowed) 1. The server requires encryption, but the client is not encrypted and the client does not support receiving error code ecNotEncrypted being returned by the server. See section 3.1.4.11.3 and section 3.2.4.1.3 for details about which client versions do not support receiving error code ecNotEncrypted. 2. The client version has been blocked by the administrator.";
                case "0x80040111":
                    return "(ecLoginFailure) Server is unable to log in user to the mailbox or public folder database.";
                case "0x000003EB":
                    return "(ecUnknownUser) The server does not recognize the szUserDN parameter as a valid enabled mailbox. For more details, see [MS-OXCSTOR] section 3.1.4.1.";
                case "0x000003F2":
                    return "(ecLoginPerm) The connection is requested for administrative access, but the authentication context associated with the binding handle does not have enough privilege.";
                case "0x80040110":
                    return "(ecVersionMismatch) The client and server versions are not compatible. The client protocol version is older than that required by the server.";
                case "0x000004E1":
                    return "(ecCachedModeRequired) The server requires the client to be running in cache mode. For details about which client versions understand this error code, see section 3.2.4.1.3.";
                case "0x000004E0":
                    return "(ecRpcHttpDisallowed) The server requires the client to not be connected via RPC/HTTP. For details about which client versions understand this error code see section 3.1.4.11.3.";
                case "0x000007D8":
                    return "(ecProtocolDisabled) The server disallows the user to access the server via this protocol interface. This could be done if the user is only capable of accessing their mailbox information through a different means (for example, Webmail, POP, or IMAP). For details about which client versions understand this error code see section 3.1.4.11.3.";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// It calls native method EcDoRpcExt2 that sends remote operation (ROP) commands to the server with a Session Context Handle.
        /// </summary>
        /// <param name="pcxh">On input, the client MUST pass a valid Session Context Handle that was created by calling EcDoConnectEx. 
        /// The server uses the Session Context Handle to identify the Session Context to use for this call. On output, the server MUST return the same Session Context Handle on success.</param>
        /// <param name="pulFlags">On input, this parameter contains flags that tell the server how to build the rgbOut parameter.</param>
        /// <param name="rgbIn">This buffer contains the ROP request payload. </param>
        /// <param name="ropRequestLength">This parameter contains the length of the ROP request payload passed in the rgbIn parameter.</param>
        /// <param name="rgbOut">On success, this buffer contains the ROP response payload.</param>
        /// <param name="pcbOut">On input, this parameter contains the maximum size of the rgbOut buffer.On output, 
        /// this parameter contains the size of the ROP response payload.</param>
        /// <param name="rgbAuxIn"> This parameter contains an auxiliary payload buffer.</param>
        /// <param name="auxInLength">On input, this parameter contains the length of the auxiliary payload buffer passed in the rgbAuxIn parameter.</param>
        /// <param name="rgbAuxOut">On output, the server can return auxiliary payload data to the client.</param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer. 
        /// On output, this parameter contains the size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="pulTransTime">On output, the server stores the number of milliseconds the call took to execute.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, return RPC error code RPC exception code.</returns>
        private uint RpcExt2(
                        ref IntPtr pcxh,
                        ref uint pulFlags,
                        byte[] rgbIn,
                        uint ropRequestLength,
                        out byte[] rgbOut,
                        ref uint pcbOut,
                        byte[] rgbAuxIn,
                        uint auxInLength,
                        out byte[] rgbAuxOut,
                        ref uint pcbAuxOut,
                        out uint pulTransTime)
        {
            this.site.Assert.AreNotEqual<IntPtr>(IntPtr.Zero, pcxh, "Session Context should not be empty.");
            this.site.Assert.AreEqual<uint>(ropRequestLength, (uint)rgbIn.Length, "Passed in buffer rgbIn length should be equal to cbIn.");
            this.site.Assert.AreEqual<uint>(auxInLength, (uint)rgbAuxIn.Length, "Passed in buffer rgbAuxIn length should be equal to cbAuxIn.");

            IntPtr rgbInPtr = Marshal.AllocHGlobal(rgbIn.Length);
            IntPtr rgbAuxInPtr = Marshal.AllocHGlobal(rgbAuxIn.Length);
            try
            {
                Marshal.Copy(rgbIn, 0, rgbInPtr, rgbIn.Length);
                Marshal.Copy(rgbAuxIn, 0, rgbAuxInPtr, rgbAuxIn.Length);
                IntPtr rgbOutPtr = Marshal.AllocHGlobal((int)pcbOut);
                IntPtr rgbAuxOutPtr = Marshal.AllocHGlobal((int)pcbAuxOut);
                uint ret = NativeMethods.EcDoRpcExt2(
                            ref pcxh,
                            ref pulFlags,
                            rgbInPtr,
                            ropRequestLength,
                            rgbOutPtr,
                            ref pcbOut,
                            rgbAuxInPtr,
                            auxInLength,
                            rgbAuxOutPtr,
                            ref pcbAuxOut,
                            out pulTransTime);

                rgbOut = new byte[(int)pcbOut];
                Marshal.Copy(rgbOutPtr, rgbOut, 0, (int)pcbOut);
                Marshal.FreeHGlobal(rgbOutPtr);

                rgbAuxOut = new byte[(int)pcbAuxOut];
                Marshal.Copy(rgbAuxOutPtr, rgbAuxOut, 0, (int)pcbAuxOut);
                Marshal.FreeHGlobal(rgbAuxOutPtr);

                if (ret != 0)
                {
                    this.site.Log.Add(LogEntryKind.Comment, "EcDoRpcExt2 returns error code {0}. Refers to [MS-OXCDATA] section 2.4 for more information.", ret);
                }

                return ret;
            }
            catch (SEHException e)
            {
                rgbOut = null;
                rgbAuxOut = null;
                pulTransTime = 0;
                uint errorCode = RpcExceptionCode(e);

                this.site.Log.Add(LogEntryKind.Comment, "EcDoRpcExt2 throws exception, system error code is {0}, the error message is: {1}", errorCode, (new Win32Exception((int)errorCode)).ToString());

                return errorCode;
            }
            finally
            {
                Marshal.FreeHGlobal(rgbInPtr);
                Marshal.FreeHGlobal(rgbAuxInPtr);
            }
        }

        /// <summary>
        /// The Connect method establishes a new Session Context with the server.
        /// </summary>
        /// <param name="pcxh">On success, the server MUST return a unique value to be used as a CXH. This unique value serves as the CXH for the client.</param>
        /// <param name="userDN">User's distinguished name (DN).</param>
        /// <param name="mapiContext">RPC call context</param>
        /// <returns>If the method succeeds, the return value is 0. 
        /// If the method fails, the return value is an implementation-specific error code or one of the protocol-defined error codes.</returns>
        private uint Connect_Internal(ref IntPtr pcxh, string userDN, ref MapiContext mapiContext)
        {
            #region Parameters for RPC connect operation.
            // For ordinary client calls this value MUST be 0x00000000.
            uint flags = 0x00000000;

            // The connection modulus is a client derived 32-bit hash value of the DN passed in field szUserDN and can be used by the server to decide 
            // which public folder replica to use when accessing public folder information when more than one replica of a folder exists. 
            // The hash can be used to distribute client access across replicas in a deterministic way for load balancing.
            uint conMod = 3741814581;

            // This field is reserved. A client MUST pass a value of 0x00000000.
            uint limit = 0;

            // The code page in which text data SHOULD be sent if Unicode format is not requested by the client on subsequent calls using this Session Context.
            uint cpid = 0x000004E4;
            if (mapiContext.CodePageId != null)
            {
                cpid = mapiContext.CodePageId.Value;
            }

            // The local ID for everything other than sorting.
            uint lcidString = 0x00000409;

            // The local ID for sorting.
            uint lcidSort = 0x00000409;

            // This value is used to link the Session Context created by this call with an existing Session Context on the server. 
            // If no session linking is requested, this value will be 0xFFFFFFFF.
            uint licxrLink = 0xFFFFFFFF;

            // The client MUST pass a value of 0x01.
            ushort canConvertCodePages = 0x01;

            // The server returns the number of milliseconds that a client SHOULD wait between polling the server for event information.
            uint cmsPollsMax = 0;

            // The server returns the number of times a client SHOULD retry future RPC calls using the CXH returned in this call. 
            // This is for client RPC calls that fail with RPC status code RPC_S_SERVER_TOO_BUSY. 
            // This is a suggested retry count for the client and SHOULD NOT be enforced by the server.
            uint retry = 0;

            // The server returns the number of milliseconds a client SHOULD wait before retrying a failed RPC call. 
            // If any future RPC call to the server using the CXH returned in this call fails with RPC status code RPC_S_SERVER_TOO_BUSY, 
            // it SHOULD wait the number of milliseconds specified in this output parameter before retrying the call. 
            // The number of times a client SHOULD retry is returned in parameter pcRetry. 
            // This is a suggested delay for the client and SHOULD NOT be enforced by the server.
            uint cmsRetryDelay = 0;

            // The server returns a session index value that is associated with the CXH returned from this call. 
            // This value in conjunction with the session creation time stamp value returned in pulTimeStamp will be passed to a subsequent EcDoConnectEx call, 
            // if the client wants to link two Session Contexts. The server MUST NOT assign two active Session Contexts the same session index value. 
            // The server is free to return any 16-bit value for the session index.
            ushort cxr = 0;

            // The server returns the distinguished name (DN) of the server.
            UIntPtr pszDNPrefix;

            // The server returns the display name of the server.
            UIntPtr pszDisplayName;

            ushort[] rgwClientVersion = new ushort[3];
            ushort[] rgwServerVersion = new ushort[3] { 0, 0, 0 };
            ushort[] rgwBestVersion = new ushort[3];

            // The client passes the client protocol version the server SHOULD use to determine what protocol functionality the client supports. 
            // For more information about how version numbers are interpreted from the wire data, see section 3.1.9.[MS-OXCMSG]
            rgwClientVersion[0] = 0x000c;
            rgwClientVersion[1] = 0x183e;
            rgwClientVersion[2] = 0x03e8;

            // On input, this parameter and parameter ulIcxrLink are used for linking the Session Context created by this call with an existing Session Context. 
            // If the ulIcxrLink parameter is not 0xFFFFFFFF, the client MUST pass in the pulTimeStamp value returned from the server on a previous call to 
            // EcDoConnectEx (see the ulIcxrLink and piCxr parameters for more details).
            uint timeStamp = 0;

            // This parameter contains an auxiliary payload buffer. The auxiliary payload buffer is prefixed by an RPC_HEADER_EXT structure. 
            // Information stored in this header determines how to interpret the data following the header. 
            // The length of the auxiliary payload buffer that includes the RPC_HEADER_EXT header is contained in parameter cbAuxIn.
            byte[] payloadBufferAuxIn = null;

            // On input, this parameter contains the length of the auxiliary payload buffer passed in the rgbAuxIn parameter. 
            // The server MUST fail with error code ecRpcFormat if the request buffer is larger than 0x00001008 bytes in size.
            uint auxinLength = 0;

            // 0x1008: Set the max size of the rgbAuxOut
            byte[] rgbAuxOut = new byte[RpcAdapter.PcbAuxOut];

            // 0x1008: Set the max size of the cbAuxOut
            uint auxOutLength = RpcAdapter.PcbAuxOut;
            #endregion

            uint returnValue = 0;
            try
            {
                // Connect to server.
                returnValue = NativeMethods.EcDoConnectEx(
                        this.bindingHandle,
                        ref pcxh,
                        userDN,
                        flags,
                        conMod,
                        limit,
                        cpid,
                        lcidString,
                        lcidSort,
                        licxrLink,
                        canConvertCodePages,
                        out cmsPollsMax,
                        out retry,
                        out cmsRetryDelay,
                        out cxr,
                        out pszDNPrefix,
                        out pszDisplayName,
                        rgwClientVersion,
                        rgwServerVersion,
                        rgwBestVersion,
                        ref timeStamp,
                        payloadBufferAuxIn,
                        auxinLength,
                        rgbAuxOut,
                        ref auxOutLength);

                if (returnValue != 0)
                {
                    this.site.Log.Add(LogEntryKind.Comment, "EcDoConnectEx returns error code {0}. Refers to [MS-OXCDATA] section 2.4 for more information.", returnValue);
                }

                Array.Copy(rgwServerVersion, mapiContext.EXServerVersion, 3);
            }
            catch (SEHException e)
            {
                returnValue = RpcExceptionCode(e);

                this.site.Log.Add(LogEntryKind.Comment, "EcDoConnectEx throws exception, system error code is {0}, the error message is: {1}", returnValue, (new Win32Exception((int)returnValue)).ToString());
            }

            return returnValue;
        }
    }
}
