namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXCRPC.
    /// </summary>
    public partial class MS_OXCRPCAdapter : ManagedAdapterBase, IMS_OXCRPCAdapter
    {
        #region Variable
        /// <summary>
        /// The RPC binding.
        /// </summary>
        private IntPtr bindingHandle = IntPtr.Zero;

        #endregion

        #region Initialize
        /// <summary>
        /// Initializes adapter
        /// </summary>
        /// <param name="testSite">The instance of the ITestSite.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCRPC";
            AdapterHelper.Initialize(testSite);
            Common.MergeConfiguration(testSite);
            this.RegisterROPDeserializer();
        }
        #endregion

        #region IMS_OXCRPCAdapter interface implementation
        /// <summary>
        /// Initializes the client and server, builds the transport tunnel between client and server.
        /// </summary>
        /// <param name="encryptionMethod">An unsigned integer indicates the authentication level for creating RPC binding</param>
        /// <param name="authnSvc">An unsigned integer indicates authentication services.</param>
        /// <param name="userName">Define user name which can be used by client to access SUT. </param>
        /// <param name="password">Define user password which can be used by client to access SUT.</param>
        /// <returns>If success, it returns true, else returns false.</returns>
        public bool InitializeRPC(uint encryptionMethod, uint authnSvc, string userName, string password)
        {
            string server = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            bool setUuid = bool.Parse(Common.GetConfigurationPropertyValue("SetUuid", this.Site));

            string spnStr = string.Empty;
            string spnFormat = Common.GetConfigurationPropertyValue("ServiceSPNFormat", this.Site);
            spnStr = Regex.Replace(spnFormat, @"\[ServerName\]", server, RegexOptions.IgnoreCase);
            
            // Create identity
            NativeMethods.CreateIdentity(domain, userName, password);

            string seqType = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(CultureInfo.InvariantCulture);
            bool rpchUseSsl = false;
            string rpchAuthScheme = string.Empty;
            if (seqType == "ncacn_http")
            {
                if (!bool.TryParse(Common.GetConfigurationPropertyValue("RpchUseSsl", this.Site), out rpchUseSsl))
                {
                    Site.Assert.Fail("Value of 'RpchUseSsl' property is {0}, which is invalid.", Common.GetConfigurationPropertyValue("RpchUseSsl", this.Site));
                }

                rpchAuthScheme = Common.GetConfigurationPropertyValue("RpchAuthScheme", this.Site).ToLower(CultureInfo.InvariantCulture);
                if (rpchAuthScheme != "basic" && rpchAuthScheme != "ntlm")
                {
                    Site.Assert.Fail("Value of 'RpchAuthScheme' property is {0}, which is invalid.", rpchAuthScheme);
                }
            }

            uint status = NativeMethods.BindToServer(server, encryptionMethod, authnSvc, seqType, rpchUseSsl, rpchAuthScheme, spnStr, null, setUuid);
            Site.Assert.AreEqual<ulong>(0, status, "The return value should be 0 if binding server successfully.");
            this.bindingHandle = NativeMethods.GetBindHandle();

            if (string.Compare(seqType, "ncacn_ip_tcp", true, CultureInfo.InvariantCulture) == 0)
            {
                if (this.bindingHandle == IntPtr.Zero)
                {
                    Site.Assert.Fail("Failed to create RPC binding handle with server over TCP : " + server);
                    return false;
                }

                #region Capture code
                this.VerifyNcacnIpTcp(this.bindingHandle);
                this.VerifyCommonRequirements(this.bindingHandle);
                #endregion Capture code
            }
            else if (string.Compare(seqType, "ncacn_http", true, CultureInfo.InvariantCulture) == 0)
            {
                if (this.bindingHandle == IntPtr.Zero)
                {
                    Site.Assert.Fail("Failed to create RPC binding handle with server over HTTP : " + server);
                    return false;
                }

                #region Capture code
                this.VerifyNcacnHttp(this.bindingHandle);
                this.VerifyCommonRequirements(this.bindingHandle);
                #endregion Capture code
            }

            return true;
        }

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
        public uint EcDoConnectEx(ref IntPtr pcxh, uint sessionContextLink, ref uint pulTimeStamp, byte[] rgbAuxIn, string userDN, ref uint pcbAuxOut, ushort[] rgwClientVersion, out ushort[] rgwServerVersion, out ushort[] rgwBestVersion, out ushort picxr, uint flags, out List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue)
        {
            if (this.bindingHandle == IntPtr.Zero)
            {
                uint encryptionMethod = (uint)Convert.ToInt32(Common.GetConfigurationPropertyValue("RpcAuthenticationLevel", this.Site));
                uint authnSvc = (uint)Convert.ToInt32(Common.GetConfigurationPropertyValue("RPCAuthenticationService", this.Site));
                string userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
                string password = Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site);
                bool returnStatus = this.InitializeRPC(encryptionMethod, authnSvc, userName, password);
                Site.Assert.IsTrue(returnStatus, "Initializing the RPC binding should success.");
            }

            picxr = 0;
            rgbAuxOutValue = null;
            uint inputAuxSize = 0;
            if (rgbAuxIn != null)
            {
                inputAuxSize = (uint)rgbAuxIn.Length;
            }

            rgwBestVersion = new ushort[3] { 0, 0, 0 };
           
            // An unsigned integer indicates the local ID for everything other than sorting, as specified in MS-OXCRPC EcDoConnectEx Method.
            uint localIdString = ConstValues.DefaultLocale;

            // An unsigned integer indicates the local ID for sorting, as specified in MS-OXCRPC EcDoConnectEx Method.
            uint localIdSort = ConstValues.DefaultLocale;

            // Must be 0 as specified in Open Specification
            uint valueOfcbLimit = 0x00000000;

            // Must be 0x01, specified in Open Specification
            ushort canConvertCodePages = 0x01;
            uint pcmsRetryDelay;
            uint pcmsPollsMax;
            uint retryTimes;
            UIntPtr valueOfDNPrefix;
            UIntPtr displayName;
            rgwServerVersion = new ushort[3] { 0, 0, 0 };
            byte[] rgbAuxOut = new byte[pcbAuxOut];
            bool isCompressed = false;
            uint returnValue = 0;

            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("ConnectRetryCount", this.Site));
            int connectCount = 0;
            do
            {
                returnValue = this.EcDoConnectEx_Internal(
                    ref pcxh,
                    userDN,
                    flags,
                    ConstValues.ConnectionMod,
                    valueOfcbLimit,
                    ConstValues.CodePageId,
                    localIdString,
                    localIdSort,
                    sessionContextLink,
                    canConvertCodePages,
                    out pcmsPollsMax,
                    out retryTimes,
                    out pcmsRetryDelay,
                    out picxr,
                    out valueOfDNPrefix,
                    out displayName,
                    rgwClientVersion,
                    rgwServerVersion,
                    rgwBestVersion,
                    ref pulTimeStamp,
                    rgbAuxIn,
                    inputAuxSize,
                    rgbAuxOut,
                    ref pcbAuxOut);

                if (returnValue >= 1700 && returnValue <= 1799 && returnValue != 0x000006F7)
                {
                    // If returnValue is between 1700 and 1799, which are about network or RPC connection, try to connect again. 
                    // The error code 0x000006F7 means "the stub received bad data". So test case should not try to connect again.
                    connectCount++;
                    System.Threading.Thread.Sleep(waitTime);

                    if (connectCount > 0)
                    {
                        Site.Log.Add(LogEntryKind.Comment, "RPC server is unavailable or the remote procedure call failed and did not execute, will try to connect the server again. Current retry number is {0}. Current returnValue is {1}.", connectCount, returnValue);
                    }
                }
                else
                {
                    break;
                }
            } 
            while (connectCount < retryCount);
            if (connectCount == retryCount)
            {
                Site.Assert.Fail("The connection is still failed after {0} times retrying.", connectCount);
            }

            #region Capture code

            // Verify method EcDoConnectEx
            this.VerifyEcDoConnectEx(
                pcxh, 
                displayName.ToString(),
                rgwClientVersion, 
                rgwServerVersion,
                rgwBestVersion,
                picxr,
                pcbAuxOut,
                pulTimeStamp,
                returnValue);

            List<short> flagValues = new List<short>();

            // Verify rgbAuxOut field
            if (pcbAuxOut != 0 && returnValue == 0)
            {
                // RPC_HEADER_EXT is in front of Payload.
                short flag = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtVersionByteSize);
                int size = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtVersionByteSize + ConstValues.RpcHeaderExtSizeByteSize) + AdapterHelper.RPCHeaderExtlength;
                int actualSize = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtSizeByteSize + ConstValues.RpcHeaderExtFlagsByteSize + ConstValues.RpcHeaderExtVersionByteSize) + AdapterHelper.RPCHeaderExtlength;
                flagValues.Add(flag);

                

                // If rgbAuxOut is obfuscated (0x02, XorMagic, means the data following the RPC_HEADER_EXT has been obfuscated)
                if ((flag & (short)RpcHeaderExtFlags.XorMagic) == (short)RpcHeaderExtFlags.XorMagic)
                {
                    byte[] rgbXorAuxOut = null;
                    byte[] rgbOriAuxOut = new byte[size];
                    Array.Copy(rgbAuxOut, 0, rgbOriAuxOut, 0, size);

                    // According to the Open Specification, every byte of the data to be obfuscated has XOR applied with the value 0xA5.
                    bool obfuscationResult = false;
                    rgbXorAuxOut = Common.XOR(rgbOriAuxOut);
                    if (rgbXorAuxOut != null)
                    {
                        obfuscationResult = true;
                    }

                    Array.Copy(rgbXorAuxOut, 0, rgbAuxOut, AdapterHelper.RPCHeaderExtlength, size - AdapterHelper.RPCHeaderExtlength);

                    // The first true means the current method is EcDoConnectEx.
                    // The second true means the current field is rgbAuxOut
                    this.VerifyObfuscationAlgorithm(obfuscationResult, true, true, flag);
                }

                // Decompress or revert if the rgbAuxOut is compressed or obfuscated.
                // If rgbAuxOut is Compressed (01, Compressed, means the data that follows the RPC_HEADER_EXT is compressed.)
                if ((flag & (short)RpcHeaderExtFlags.Compressed) == (short)RpcHeaderExtFlags.Compressed)
                {
                    isCompressed = true;
                    byte[] rgbOriAuxOut = new byte[size];
                    Array.Copy(rgbAuxOut, 0, rgbOriAuxOut, 0, size);
                    bool decompressResult = false;
                    try
                    {
                        rgbAuxOut = Common.DecompressStream(rgbOriAuxOut);
                        decompressResult = true;
                    }
                    catch (ArgumentException)
                    {
                        decompressResult = false;
                    }
                    catch (InvalidOperationException)
                    {
                        decompressResult = false;
                    }

                    // The first true means the current method is EcDoConnectEx.
                    // The second true means the current field is rgbAuxOut
                    this.VerifyCompressionAlgorithm(decompressResult, true, true, flag);
                    this.VerifyDIRECT2EncodingAlgorithm(decompressResult);
                }

                if (rgbAuxOut != null)
                {
                    if (isCompressed)
                    {
                        Array.Resize<byte>(ref rgbAuxOut, actualSize);
                    }
                    else
                    {
                        Array.Resize<byte>(ref rgbAuxOut, (int)pcbAuxOut);
                    }
                }

                #region Verify RPC header
                ExtendedBuffer[] buffer = ExtractExtendedBuffer(rgbAuxOut);
                this.VerifyRPCHeaderExt(flagValues, buffer[0].Header, true, true, true);
                #endregion
                rgbAuxOutValue = this.ParseRgbAuxOut(rgbAuxOut);
                this.VerifyRgbAuxOutPayLoadOnEcDoConnectEx(rgbAuxOutValue);
            }
            #endregion Capture code

            return returnValue;
        }

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
        public uint EcDoConnectEx(ref IntPtr pcxh, uint sessionContextLink, ref uint pulTimeStamp, byte[] rgbAuxIn, string userDN, ref uint pcbAuxOut, ushort[] rgwClientVersion, out ushort[] rgwBestVersion, out ushort picxr)
        {
            ushort[] rgwServerVersion;

            // Administrator privilege requested for connection specified in the Open Specification.
            uint flags = 1;
            return this.EcDoConnectEx(ref pcxh, sessionContextLink, ref pulTimeStamp, rgbAuxIn, userDN, ref pcbAuxOut, rgwClientVersion, out rgwServerVersion, out rgwBestVersion, out picxr, flags);
        }

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
        public uint EcDoConnectEx(ref IntPtr pcxh, uint sessionContextLink, ref uint pulTimeStamp, byte[] rgbAuxIn, string userDN, ref uint pcbAuxOut, ushort[] rgwClientVersion, out ushort[] rgwServerVersion, out ushort[] rgwBestVersion, out ushort picxr, uint flags)
        {
            List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue;
            return this.EcDoConnectEx(ref pcxh, sessionContextLink, ref pulTimeStamp, rgbAuxIn, userDN, ref pcbAuxOut, rgwClientVersion, out rgwServerVersion, out rgwBestVersion, out picxr, flags, out rgbAuxOutValue);
        }

        /// <summary>
        /// The method EcRRegisterPushNotification registers a callback address with the server for a Session Context. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="rgbContext">This parameter contains opaque client-generated context data that is sent back to the client at the callback address.</param>
        /// <param name="addType">The type of the cbCallbackAddress.</param>
        /// <param name="ip">The client IP used in this method.</param>
        /// <param name="notificationHandle">If the call completes successfully, this output parameter will contain a handle to the notification callback on the server.</param>
        /// <returns>If success, it returns 0, else returns the error code.</returns>
        public uint EcRRegisterPushNotification(ref IntPtr pcxh, byte[] rgbContext, Add_Families addType, string ip, out uint notificationHandle)
        {
            notificationHandle = 0;
            uint retValue = 0;
            ushort clientContextSize = (ushort)(rgbContext == null ? 0 : rgbContext.Length);

            try
            {
                ushort port = ushort.Parse(Common.GetConfigurationPropertyValue("NotificationPort", this.Site)); 
                IntPtr oldPcxh = pcxh;
                retValue = NativeMethods.EcRRegisterPushNotificationWrap(
                    ref pcxh,
                    (short)addType,
                    ip,
                    port,
                    rgbContext,
                    clientContextSize,
                    out notificationHandle);
                this.VerifyEcRRegisterPushNotification(pcxh, retValue, oldPcxh);
            }
            catch (SEHException e)
            {
                retValue = NativeMethods.RpcExceptionCode(e);
            }

            return retValue;
        }

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
        public uint EcDoRpcExt2(
            ref IntPtr pcxh,
            PulFlags pulFlags, 
            byte[] rgbIn,
            ref uint pcbOut,
            byte[] rgbAuxIn,
            ref uint pcbAuxOut,
            out IDeserializable response,
            ref List<List<uint>> responseSOHTable)
        {
            byte[] rgbOut = new byte[pcbOut];
            byte[] rgbAuxOut = new byte[pcbAuxOut];
            uint payloadCount = 0;
            return this.EcDoRpcExt2(ref pcxh, pulFlags, rgbIn, ref rgbOut, ref pcbOut, rgbAuxIn, ref pcbAuxOut, out response, ref responseSOHTable, out payloadCount, ref rgbAuxOut);
        }

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
        /// <param name="rgbAuxOut">The data of the output buffer rgbAuxOut. </param>
        /// <returns>If success, it return 0, else return the error code.</returns>
        public uint EcDoRpcExt2(
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
            ref byte[] rgbAuxOut)
        {
            #region Variables
            uint inputPcbOutValue = pcbOut;
            uint inputPcbAuxOutValue = pcbAuxOut;
            IntPtr pcxhInput = pcxh;
            int pcxhValue = pcxh.ToInt32();
            uint flags = (uint)pulFlags;
            payloadCount = 0;
            uint returnValue = 0;
            uint pulTransTime;
            uint inputROPSize = 0;
            if (rgbIn != null)
            {
                inputROPSize = (uint)rgbIn.Length;
            }
            else
            {
                rgbIn = new byte[0];
            }

            uint inputAuxSize = 0;
            if (rgbAuxIn != null)
            {
                inputAuxSize = (uint)rgbAuxIn.Length;
            }
            else
            {
                rgbAuxIn = new byte[0];
            }

            rgbOut = new byte[pcbOut];
            response = null;
            #endregion

            IntPtr rgbInPtr = Marshal.AllocHGlobal(rgbIn.Length);
            IntPtr rgbAuxInPtr = Marshal.AllocHGlobal(rgbAuxIn.Length);

            try
            {
                Marshal.Copy(rgbIn, 0, rgbInPtr, rgbIn.Length);
                Marshal.Copy(rgbAuxIn, 0, rgbAuxInPtr, rgbAuxIn.Length);

                // Record whether the server is busy and needs to retry later.
                bool needRetry = false;
                int retryCount = 0;
                int maxRetryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                do
                {
                    IntPtr rgbOutPtr = Marshal.AllocHGlobal((int)pcbOut);
                    IntPtr rgbAuxOutPtr = Marshal.AllocHGlobal((int)pcbAuxOut);

                    needRetry = false;
                    returnValue = NativeMethods.EcDoRpcExt2(
                        ref pcxhInput,
                        ref flags,
                        rgbInPtr,
                        inputROPSize,
                        rgbOutPtr,
                        ref pcbOut,
                        rgbAuxInPtr,
                        inputAuxSize,
                        rgbAuxOutPtr,
                        ref pcbAuxOut,
                        out pulTransTime);

                    rgbOut = new byte[(int)pcbOut];
                    Marshal.Copy(rgbOutPtr, rgbOut, 0, (int)pcbOut);
                    Marshal.FreeHGlobal(rgbOutPtr);

                    rgbAuxOut = new byte[(int)pcbAuxOut];
                    Marshal.Copy(rgbAuxOutPtr, rgbAuxOut, 0, (int)pcbAuxOut);
                    Marshal.FreeHGlobal(rgbAuxOutPtr);

                    #region Verify requirements while EcDoRpcExt2 success

                    if (returnValue == 0 && pcbOut > 0)
                    {
                        RPC_HEADER_EXT[] rpcHeaderExts;
                        byte[][] rops;
                        uint[][] serverHandleObjectsTables;

                        byte[] rawData = new byte[pcbOut];
                        Array.Copy(rgbOut, rawData, pcbOut);
                        int size;
                        int actualSize;
                        int payLoadSizeCount = 0;
                        int payLoadActualSizeCount = 0;
                        short flag;
                        bool isCompressed = false;
                        List<byte> actualData = new List<byte>();
                        List<short> flagValuesBeforeDecompressing = new List<short>();

                        #region Decompress or revert rgbOut if it is compressed or obfuscated
                        do
                        {
                            flag = BitConverter.ToInt16(rawData, payLoadSizeCount + ConstValues.RpcHeaderExtVersionByteSize);
                            flagValuesBeforeDecompressing.Add(flag);
                            size = BitConverter.ToInt16(rawData, payLoadSizeCount + ConstValues.RpcHeaderExtVersionByteSize + ConstValues.RpcHeaderExtFlagsByteSize);
                            actualSize = BitConverter.ToInt16(rawData, payLoadSizeCount + ConstValues.RpcHeaderExtVersionByteSize + ConstValues.RpcHeaderExtFlagsByteSize + ConstValues.RpcHeaderExtSizeByteSize);
                            byte[] resultByte = new byte[actualSize + AdapterHelper.RPCHeaderExtlength];
                            Array.Copy(rawData, payLoadSizeCount, resultByte, 0, size + AdapterHelper.RPCHeaderExtlength);

                            // Compressed (0x00000001) means the data that follows the RPC_HEADER_EXT is compressed. 
                            if ((flag & (short)RpcHeaderExtFlags.Compressed) == (short)RpcHeaderExtFlags.Compressed)
                            {
                                isCompressed = true;
                                byte[] rgbOriOut = new byte[size + AdapterHelper.RPCHeaderExtlength];
                                Array.Copy(rawData, payLoadSizeCount, rgbOriOut, 0, size + AdapterHelper.RPCHeaderExtlength);
                                bool decompressResult = false;
                                try
                                {
                                    resultByte = Common.DecompressStream(rgbOriOut);
                                    decompressResult = true;
                                }
                                catch (ArgumentException)
                                {
                                    decompressResult = false;
                                }
                                catch (InvalidOperationException)
                                {
                                    decompressResult = false;
                                }

                                this.VerifyCompressionAlgorithm(decompressResult, false, false, flag);
                                this.VerifyDIRECT2EncodingAlgorithm(decompressResult);
                            }

                            // XorMagic(0x00000002) means that the data follows the RPC_HEADER_EXT has been obfuscated.
                            if ((flag & (short)RpcHeaderExtFlags.XorMagic) == (short)RpcHeaderExtFlags.XorMagic)
                            {
                                byte[] rgbXorOut = null;
                                byte[] rgbOriOut = new byte[size + AdapterHelper.RPCHeaderExtlength];
                                Array.Copy(rawData, payLoadSizeCount, rgbOriOut, 0, size + AdapterHelper.RPCHeaderExtlength);

                                // According to the Open Specification, every byte of the data to be obfuscated has XOR applied with the value 0xA5.
                                bool obfuscationResult = false;
                                rgbXorOut = Common.XOR(rgbOriOut);
                                if (rgbXorOut != null)
                                {
                                    obfuscationResult = true;
                                }

                                Array.Copy(rawData, payLoadSizeCount, resultByte, 0, AdapterHelper.RPCHeaderExtlength);
                                Array.Copy(rgbXorOut, 8, resultByte, AdapterHelper.RPCHeaderExtlength, actualSize);

                                // The first false means current method is EcDoRpcExt2
                                // The second false means current field is rgbOut
                                this.VerifyObfuscationAlgorithm(obfuscationResult, false, false, flag);
                            }

                            actualData.AddRange(resultByte);
                            payLoadSizeCount += size + AdapterHelper.RPCHeaderExtlength;
                            payLoadActualSizeCount += actualSize + AdapterHelper.RPCHeaderExtlength;
                        }
                        while ((flag & (short)RpcHeaderExtFlags.Last) != (short)RpcHeaderExtFlags.Last);
                        rgbOut = actualData.ToArray();
                        #endregion

                        #region Resize rgbOut
                        if (rgbOut != null)
                        {
                            if (isCompressed)
                            {
                                Array.Resize<byte>(ref rgbOut, payLoadActualSizeCount + 1);
                            }
                            else
                            {
                                Array.Resize<byte>(ref rgbOut, (int)pcbOut);
                            }
                        }
                        #endregion

                        this.ParseResponseBuffer(rgbOut, out rpcHeaderExts, out rops, out serverHandleObjectsTables);

                        #region Verify RPC_HEADER_EXTs in rgbOut buffer

                        for (int i = 0; i < rpcHeaderExts.Length; i++)
                        {
                            this.VerifyRPCHeaderExt(flagValuesBeforeDecompressing, rpcHeaderExts[i], false, false, i == rpcHeaderExts.Length - 1);
                        }

                        this.VerifyMultiRPCHeader(rpcHeaderExts.Length);

                        #endregion

                        #region Deserialize Rops

                        List<IDeserializable> responseRops = new List<IDeserializable>();

                        int ropID = 0;
                        if (rops != null)
                        {
                            // Verify the length of payload in the rgbOut
                            this.VerifyPayloadLengthResponse(rops);

                            // Only contains one ROP response
                            if (rops.Length == 1)
                            {
                                ropID = rops[0][0];
                                if (ropID == (int)RopId.RopBackoff)
                                {
                                    needRetry = true;

                                    // Revert the value of variables whose type is a reference type and used by method EcDoRpcExt2.
                                    pcxh = pcxhInput;
                                    flags = (uint)pulFlags;
                                    pcbOut = inputPcbOutValue;
                                    pcbAuxOut = inputPcbAuxOutValue;

                                    // Wait a period of time before retrying since the server is busy now.
                                    System.Threading.Thread.Sleep(int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site)));
                                    retryCount++;
                                    continue;
                                }
                                
                                IDeserializable ropRes = null;
                                RopDeserializer.Deserialize(rops[0], ref ropRes);
                                payloadCount = 1;

                                // True means RopDeserializer.Deserialize method is successful(Because there is no exception thrown and code can arrive here).
                                // 1 means one ROP response
                                this.VerifyIsRopResponse(true, 1);
                                responseRops.Add(ropRes);
                            }
                            else
                            {
                                // Contains more than one ROP response
                                ropID = rops[0][0];
                                foreach (byte[] rop in rops)
                                {
                                    if (rop[0] == (byte)RopId.RopBackoff)
                                    {
                                        needRetry = true;

                                        // Revert the value of variables whose type is a reference type and used by method EcDoRpcExt2.
                                        pcxh = pcxhInput;
                                        flags = (uint)pulFlags;
                                        pcbOut = inputPcbOutValue;
                                        pcbAuxOut = inputPcbAuxOutValue;

                                        // Wait a period of time before retrying since the server is busy now.
                                        System.Threading.Thread.Sleep(int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site)));
                                        retryCount++;
                                        break;
                                    }

                                    IDeserializable ropRes = null;
                                    RopDeserializer.Deserialize(rop, ref ropRes);
                                    payloadCount = (uint)rops.Length;

                                    // True means RopDeserializer.Deserialize method is successful(Because there is no exception thrown and code can arrive here).
                                    this.VerifyIsRopResponse(true, rops.Length);
                                    switch (rop[0])
                                    {
                                        // RopReadStream Response Buffer.
                                        case (int)RopId.RopReadStream:
                                            RopReadStreamResponse readStreamResponse = (RopReadStreamResponse)ropRes;
                                            responseRops.Add(readStreamResponse);
                                            break;

                                        // RopQueryRows Response Buffer.
                                        case (int)RopId.RopQueryRows:
                                            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)ropRes;
                                            responseRops.Add(queryRowsResponse);
                                            break;

                                        // RopFastTransferSourceGetBuffer Response Buffer.
                                        case (int)RopId.RopFastTransferSourceGetBuffer:
                                            RopFastTransferSourceGetBufferResponse bufferResponse = (RopFastTransferSourceGetBufferResponse)ropRes;
                                            responseRops.Add(bufferResponse);
                                            break;
                                        default:
                                            throw new InvalidCastException("The returned ROP ID is {0}.", ropID);
                                    }
                                }

                                switch (ropID)
                                {
                                    // RopReadStream Response Buffer.
                                    case (int)RopId.RopReadStream:
                                        for (int i = 0; i < responseRops.Count; i++)
                                        {
                                            for (int j = i + 1; j < responseRops.Count; j++)
                                            {
                                                RopReadStreamResponse readStreamResponse0 = (RopReadStreamResponse)responseRops[i];
                                                RopReadStreamResponse readStreamResponse1 = (RopReadStreamResponse)responseRops[j];
                                                this.VerifyRopReadStreamResponse(readStreamResponse0.DataSize, readStreamResponse1.DataSize, responseRops.Count);
                                            }
                                        }

                                        break;

                                    // RopQueryRows Response Buffer.
                                    case (int)RopId.RopQueryRows:
                                        for (int i = 0; i < responseRops.Count; i++)
                                        {
                                            for (int j = i + 1; j < responseRops.Count; j++)
                                            {
                                                RopQueryRowsResponse queryRowsResponse0 = (RopQueryRowsResponse)responseRops[i];
                                                RopQueryRowsResponse queryRowsResponse1 = (RopQueryRowsResponse)responseRops[j];
                                                this.VerifyRopQueryRowsResponse(queryRowsResponse0.RowCount, queryRowsResponse1.RowCount, responseRops.Count);
                                            }
                                        }

                                        break;

                                    // RopFastTransferSourceGetBuffer Response Buffer.
                                    case (int)RopId.RopFastTransferSourceGetBuffer:
                                        this.VerifyRopFastTransferSourceGetBufferResponse(responseRops.Count);
                                        break;
                                    default:
                                        throw new InvalidCastException("The returned ROP ID is {0}.", ropID);
                                }
                            }

                            if (responseRops != null)
                            {
                                if (responseRops.Count > 0)
                                {
                                    response = responseRops[0];
                                }
                            }
                            else
                            {
                                response = null;
                            }
                        }

                        #endregion

                        #region Deserialize SOH

                        if (needRetry)
                        {
                            continue;
                        }

                        responseSOHTable = new List<List<uint>>();

                        // Deserialize serverHandleObjectsTables
                        if (serverHandleObjectsTables != null)
                        {
                            foreach (uint[] serverHandleObjectsTable in serverHandleObjectsTables)
                            {
                                List<uint> serverHandleObjectList = new List<uint>();
                                foreach (uint serverHandleObject in serverHandleObjectsTable)
                                {
                                    serverHandleObjectList.Add(serverHandleObject);
                                }

                                responseSOHTable.Add(serverHandleObjectList);
                            }
                        }

                        #endregion

                        #region Parse rgbAuxOut

                        #region Capture code

                        #region Verify rgbAuxOut
                        this.VerifyIsRgbSupportedOnEcDoRpcExt2(pcbAuxOut);

                        List<short> flagValues = new List<short>();

                        if (pcbAuxOut != 0)
                        {
                            isCompressed = false;

                            // Decompress or revert if the rgbAuxOut is compressed or obfuscated
                            flag = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtVersionByteSize);
                            size = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtFlagsByteSize + ConstValues.RpcHeaderExtSizeByteSize) + AdapterHelper.RPCHeaderExtlength;
                            actualSize = BitConverter.ToInt16(rgbAuxOut, ConstValues.RpcHeaderExtSizeByteSize + ConstValues.RpcHeaderExtFlagsByteSize + ConstValues.RpcHeaderExtVersionByteSize) + AdapterHelper.RPCHeaderExtlength;
                            flagValues.Add(flag);

                            // Decompress or revert rgbAuxOut if it is compressed or obfuscated
                            // Compressed (0x00000001) means the data that follows the RPC_HEADER_EXT is compressed.
                            if ((flag & (short)RpcHeaderExtFlags.Compressed) == (short)RpcHeaderExtFlags.Compressed)
                            {
                                isCompressed = true;
                                byte[] rgbOriAuxOut = new byte[size];
                                Array.Copy(rgbAuxOut, 0, rgbOriAuxOut, 0, size);
                                bool decompressResult = false;
                                try
                                {
                                    rgbAuxOut = Common.DecompressStream(rgbOriAuxOut);
                                    decompressResult = true;
                                }
                                catch (ArgumentException)
                                {
                                    decompressResult = false;
                                }
                                catch (InvalidOperationException)
                                {
                                    decompressResult = false;
                                }

                                // False means current method is EcDoRpcExt2
                                // True means current field is rgbAuxOut
                                this.VerifyCompressionAlgorithm(decompressResult, false, true, flag);
                                this.VerifyDIRECT2EncodingAlgorithm(decompressResult);
                            }

                            // XorMagic(0x00000002) means that the data follows the RPC_HEADER_EXT has been obfuscated.
                            if ((flag & (short)RpcHeaderExtFlags.XorMagic) == (short)RpcHeaderExtFlags.XorMagic)
                            {
                                byte[] rgbXorAuxOut = null;
                                byte[] rgbOriAuxOut = new byte[size];
                                Array.Copy(rgbAuxOut, AdapterHelper.RPCHeaderExtlength, rgbOriAuxOut, 0, size);

                                // Every byte of the data to be obfuscated has XOR applied with the value 0xA5.
                                bool obfuscationResult = false;
                                rgbXorAuxOut = Common.XOR(rgbOriAuxOut);
                                if (rgbXorAuxOut != null)
                                {
                                    obfuscationResult = true;
                                }

                                Array.Copy(rgbXorAuxOut, 0, rgbAuxOut, AdapterHelper.RPCHeaderExtlength, size - AdapterHelper.RPCHeaderExtlength);
                                this.VerifyObfuscationAlgorithm(obfuscationResult, false, true, flag);
                            }

                            if (rgbAuxOut != null)
                            {
                                if (isCompressed)
                                {
                                    Array.Resize<byte>(ref rgbAuxOut, actualSize);
                                }
                                else
                                {
                                    Array.Resize<byte>(ref rgbAuxOut, (int)pcbAuxOut);
                                }
                            }
                            #region Verify RPC header
                            ExtendedBuffer[] buffer = ExtractExtendedBuffer(rgbAuxOut);
                            this.VerifyRPCHeaderExt(flagValues, buffer[0].Header, false, true, true);
                            #endregion
                            List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue = this.ParseRgbAuxOut(rgbAuxOut);
                            this.VerifyRgbAuxOutOnEcDoRpcExt2(rgbAuxOutValue);
                        }
                        #endregion

                        #endregion Capture code

                        #endregion
                    }

                    // Verify method EcDoRpcExt2.
                    this.VerifyEcDoRpcExt2(pcxhValue, pcxhInput.ToInt32(), pcbOut, pcbAuxOut, inputPcbOutValue, rgbOut, inputPcbAuxOutValue, rgbAuxOut, returnValue, (uint)pulFlags, flags);
                    #endregion
                }
                while (needRetry && retryCount < maxRetryCount);
                if (needRetry && retryCount == maxRetryCount)
                {
                    Site.Assert.Fail("Server is still busy after retrying {0} times which is defined in the common configuration file and the returned ROP is RopBackOff.", retryCount);
                }
            }
            catch (SEHException e)
            {
                // Uses try...catch... code structure to get the error code of RPC call
                returnValue = NativeMethods.RpcExceptionCode(e);

                Site.Log.Add(LogEntryKind.Comment, "EcDoRpcExt2 throws exception, system error code is {0}, the error message is: {1}", returnValue, (new Win32Exception((int)returnValue)).ToString());
            }
            finally
            {
                Marshal.FreeHGlobal(rgbInPtr);
                Marshal.FreeHGlobal(rgbAuxInPtr);
            }

            return returnValue;
        }

        /// <summary>
        /// The method EcDoDisconnect closes the Session Context with the server. 
        /// </summary>
        /// <param name="pcxh">The unique value points to a CXH.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        public uint EcDoDisconnect(ref IntPtr pcxh)
        {
            uint returnValue = 0;
            try
            {
                returnValue = NativeMethods.EcDoDisconnect(ref pcxh);

                if (this.bindingHandle != IntPtr.Zero && bool.Parse(Common.GetConfigurationPropertyValue("RpcForceShutdownAssociation", this.Site)))
                {
                    uint status = NativeMethods.RpcBindingSetOption(this.bindingHandle, 13, 1); // 13 represents RPC_C_OPT_DONT_LINGER option
                    if (status != 0)
                    {
                        throw new Exception("Fails to set option on the binding handle.");
                    }
                }

                if (this.bindingHandle != IntPtr.Zero)
                {
                    NativeMethods.RpcBindingFree(ref this.bindingHandle);
                    this.bindingHandle = IntPtr.Zero;
                }
            }
            catch (SEHException e)
            {
                // Uses try...catch... code structure to get the error code of RPC call
                returnValue = NativeMethods.RpcExceptionCode(e);
            }

            #region Capture code
            this.VerifyEcDoDisconnect(pcxh.ToInt32(), returnValue);
            #endregion Capture code
            return returnValue;
        }

        /// <summary>
        /// The method EcDoAsyncConnectEx binds a Session Context Handle (CXH) returned from method EcDoConnectEx to a new Asynchronous Context Handle (ACXH) 
        /// that can be used in calls to EcDoAsyncWaitEx in interface AsyncEMSMDB. 
        /// </summary>
        /// <param name="pcxh">A unique value to be used as a CXH.</param>
        /// <param name="pacxh">An ACXH that is associated with the Session Context passed in parameter CXH.</param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        public uint EcDoAsyncConnectEx(IntPtr pcxh, ref IntPtr pacxh)
        {
            uint retValue = 0;
            try
            {
                retValue = NativeMethods.EcDoAsyncConnectEx(pcxh, ref pacxh);
                this.VerifyEcDoAsyncConnectEx(retValue, pacxh);
            }
            catch (SEHException e)
            {
                retValue = NativeMethods.RpcExceptionCode(e);
            }
                                                                                                                                                                                                
            return retValue;
        }

        /// <summary>
        /// This RPC method determines if it can communicate with the server.
        /// </summary>
        /// <returns>Always return zero.</returns>
        public uint EcDummyRpc()
        {
            uint retValue = 0;
            try
            {
                retValue = NativeMethods.EcDummyRpc(this.bindingHandle);
            }
            catch (SEHException e)
            {
                retValue = NativeMethods.RpcExceptionCode(e);
            }
            #region Capture code
            this.VerifyEcDummyRpc(retValue);
            #endregion Capture code
            return retValue;
        }

        /// <summary>
        /// The method EcDoAsyncWaitEx is an asynchronous call that will not be completed by server until there are pending events on the Session Context up to a five minute duration. 
        /// </summary>
        /// <param name="acxh">The unique value to be used as a CXH.</param>
        /// <param name="isNotificationPending">A Boolean value indicates signals that events are pending for the client on the Session Context on the server. </param>
        /// <returns>If success, it returns 0, else returns the error code specified in the Open Specification.</returns>
        public uint EcDoAsyncWaitEx(IntPtr acxh, out bool isNotificationPending)
        {
            isNotificationPending = true;
            uint pulFlagsOut = 0;
            uint retValue = 0;
            uint maxWaitTime = (uint)Convert.ToInt32(Common.GetConfigurationPropertyValue("MaxWaitTime", this.Site));

            IntPtr rpcAsyncHandle = NativeMethods.CreateRpcAsyncHandle();
            Site.Assert.AreNotEqual<IntPtr>(IntPtr.Zero, rpcAsyncHandle, "Server should return valid asynchronous handle.");
            IntPtr pulFlagsOutPtr = Marshal.AllocHGlobal(sizeof(uint));
            try
            {
                NativeMethods.EcDoAsyncWaitEx(rpcAsyncHandle, acxh, 0, pulFlagsOutPtr);

                uint waitTime = 0;
                RPCAsyncStatus getCallstatus;
                do
                {
                    // Get the status of EcDoAsyncWaitEx call.
                    // The event has not been triggered yet, so the call should not be completed.
                    getCallstatus = (RPCAsyncStatus)NativeMethods.RpcAsyncGetCallStatus(rpcAsyncHandle);
                    if (getCallstatus != RPCAsyncStatus.RPC_S_ASYNC_CALL_PENDING)
                    {
                        break;
                    }

                    System.Threading.Thread.Sleep(1000);
                    waitTime++;
                }
                while (waitTime < maxWaitTime);
                Site.Assert.AreEqual<RPCAsyncStatus>(RPCAsyncStatus.RPC_S_OK, getCallstatus, "Invoking RpcAsyncGetCallStatus method should be successful.");

                IntPtr reply = Marshal.AllocHGlobal(sizeof(int));
                RPCAsyncStatus callCompleteStatus = (RPCAsyncStatus)NativeMethods.RpcAsyncCompleteCall(rpcAsyncHandle, reply);
                Site.Assert.AreEqual<RPCAsyncStatus>(RPCAsyncStatus.RPC_S_OK, callCompleteStatus, "Invoking RpcAsyncCompleteCall method for completing asynchronous wait should be successful.");
                retValue = (uint)Marshal.ReadInt32(reply);
                pulFlagsOut = (uint)Marshal.ReadInt32(pulFlagsOutPtr);
            }
            catch (SEHException e)
            {
                retValue = NativeMethods.RpcExceptionCode(e);
            }
                                        
            isNotificationPending = pulFlagsOut == 0 ? false : true;

            if (retValue == 0 && pulFlagsOut != 0)
            {
                this.VerifyEcDoAsyncWaitExpulFlagsOut(pulFlagsOut);
            }
                                                                                                                                                                                                
            this.VerifyEcDoAsyncWaitEx(acxh, retValue);
            return retValue;
        }

        #endregion

        #region Helps
        /// <summary>
        /// Extracts RPC_HEADER_EXT structure from rgbAuxOut buffer.
        /// </summary>
        /// <param name="payload">The rgbAuxOut buffer contained in the response buffer.</param>
        /// <param name="headerExt">The RPC_HEADER_EXT structure. </param>
        /// <returns>The flag to indicate whether the rgbAuxOut is a valid buffer.</returns>
        private static bool ExtractExtendedHeader(byte[] payload, out RPC_HEADER_EXT headerExt)
        {
            headerExt = new RPC_HEADER_EXT();
            if (payload.Length < AdapterHelper.RPCHeaderExtlength)
            {
                return false;
            }

            headerExt.Version = BitConverter.ToUInt16(payload, 0);
            headerExt.Flags = BitConverter.ToUInt16(payload, ConstValues.RpcHeaderExtSizeByteSize);
            headerExt.Size = BitConverter.ToUInt16(payload, ConstValues.RpcHeaderExtVersionByteSize + ConstValues.RpcHeaderExtFlagsByteSize);
            headerExt.SizeActual = BitConverter.ToUInt16(payload, ConstValues.RpcHeaderExtVersionByteSize + ConstValues.RpcHeaderExtFlagsByteSize + ConstValues.RpcHeaderExtSizeByteSize);
            return true;
        }

        /// <summary>
        /// Extracts the payloads contained in the response buffer.
        /// </summary>
        /// <param name="response">The rgbOut or rgbAuxOut buffer returned from server.</param>
        /// <returns>The payloads array that was extracted from the response buffer.</returns>
        private static ExtendedBuffer[] ExtractExtendedBuffer(byte[] response)
        {
            if (response == null)
            {
                return null;
            }

            if (response.Length == 0)
            {
                return null;
            }

            bool errorBuffer = false;
            for (int i = 0; i < response.Length; i++)
            {
                if (response[i] != 0)
                {
                    errorBuffer = true;
                }
            }

            if (!errorBuffer)
            {
                return null;
            }

            int pos = 0;
            int index = 0;

            ExtendedBuffer[] buffer = new ExtendedBuffer[ConstValues.ArbitraryInitialSizeForBuffer];
            while (pos + AdapterHelper.RPCHeaderExtlength < response.Length + 1)
            {
                byte[] hdr_Bytes = new byte[AdapterHelper.RPCHeaderExtlength];
                System.Array.Copy(response, pos, hdr_Bytes, 0, hdr_Bytes.Length);

                RPC_HEADER_EXT hdr;
                if (!ExtractExtendedHeader(hdr_Bytes, out hdr))
                {
                    break;
                }
                else if (pos + AdapterHelper.RPCHeaderExtlength + hdr.SizeActual > response.Length)
                {
                    break;
                }
                else
                {
                    buffer[index].Header = hdr;
                    buffer[index].Payload = new byte[hdr.SizeActual];
                    System.Array.Copy(response, pos + AdapterHelper.RPCHeaderExtlength, buffer[index].Payload, 0, hdr.SizeActual);
                    index++;
                    pos += AdapterHelper.RPCHeaderExtlength + hdr.SizeActual;
                }
            }

            System.Array.Resize<ExtendedBuffer>(ref buffer, index);
            return buffer;
        }

        /// <summary>
        /// The EcDoConnectEx_Internal method establishes a new Session Context with the server.
        /// </summary>
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
        /// <param name="rgwClientVersion">The client passes the client protocol version the server SHOULD use to determine what protocol functionality the client supports.</param>
        /// <param name="rgwServerVersion">The server returns the server protocol version the client SHOULD use to determine what protocol functionality the server supports.</param>
        /// <param name="rgwBestVersion">The server returns the minimum client protocol version the server supports. This information is useful if the EcDoConnectEx call fails with return code ecVersionMismatch. On success, the server SHOULD return the value passed in rgwClientVersion by the client. For details about how version numbers are interpreted from the wire data.</param>
        /// <param name="pulTimeStamp">On input, this parameter and parameter ulIcxrLink are used for linking the Session Context created by this call with an existing Session Context. If the ulIcxrLink parameter is not 0xFFFFFFFF, the client MUST pass in the pulTimeStamp value returned from the server on a previous call to EcDoConnectEx (see the ulIcxrLink and piCxr parameters for more details).</param>
        /// <param name="rgbAuxIn">This parameter contains an auxiliary payload buffer. The auxiliary payload buffer is prefixed by an RPC_HEADER_EXT structure. Information stored in this header determines how to interpret the data following the header. The length of the auxiliary payload buffer that includes the RPC_HEADER_EXT header is contained in parameter cbAuxIn.</param>
        /// <param name="inputAuxSize">On input, this parameter contains the length of the auxiliary payload buffer passed in the rgbAuxIn parameter. The server MUST fail with error code ecRpcFormat if the request buffer is larger than 0x00001008 bytes in size.</param>
        /// <param name="rgbAuxOut">On output, the server can return auxiliary payload data to the client. The server MUST include an RPC_HEADER_EXT header before the auxiliary payload data.</param>
        /// <param name="pcbAuxOut">On input, this parameter contains the maximum length of the rgbAuxOut buffer. The server MUST fail with error code ecRpcFormat if this value is larger than 0x00001008. On output, this parameter contains the size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code or one of the protocol-defined error codes. No exceptions are thrown beyond those thrown by the underlying RPC protocol [MS-RPCE].</returns>
        private uint EcDoConnectEx_Internal(
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
            ref uint pcbAuxOut)
        {
            uint returnValue;
            pcmsPollsMax = 0;
            retryTimes = 0;
            valueOfDNPrefix = UIntPtr.Zero;
            displayName = UIntPtr.Zero;
            picxr = 0;
            pcmsRetryDelay = 0;

            try
            {
                returnValue = NativeMethods.EcDoConnectEx(
                    this.bindingHandle,
                    ref pcxh,
                    userDN,
                    flags,
                    connectionModulus,
                    limit,
                    codePageId,
                    localIdString,
                    localIdSort,
                    sessionContextLink,
                    isCanConvertCodePages,
                    out pcmsPollsMax,
                    out retryTimes,
                    out pcmsRetryDelay,
                    out picxr,
                    out valueOfDNPrefix,
                    out displayName,
                    rgwClientVersion,
                    rgwServerVersion,
                    rgwBestVersion,
                    ref pulTimeStamp,
                    rgbAuxIn,
                    inputAuxSize,
                    rgbAuxOut,
                    ref pcbAuxOut);

                if (returnValue != 0)
                {
                    Site.Log.Add(LogEntryKind.Comment, "EcDoConnectEx returns error code {0}. Refers to [MS-OXCDATA] section 2.4 for more information.", returnValue);
                }
            }
            catch (SEHException e)
            {
                // Uses try...catch... code structure to get the error code of RPC call
                returnValue = NativeMethods.RpcExceptionCode(e);

                Site.Log.Add(LogEntryKind.Comment, "EcDoConnectEx throws exception, system error code is {0}, the error message is: {1}", returnValue, (new Win32Exception((int)returnValue)).ToString());
            }
        
            return returnValue;
        }

        /// <summary>
        /// The method parses ROP response buffer.
        /// </summary>
        /// <param name="rgbOut">The ROP response payload.</param>
        /// <param name="rpcHeaderExts">Array of RPCIHEADER_EXT.</param>
        /// <param name="rops">Byte array contains parsed ROP command.</param>
        /// <param name="serverHandleObjectsTables">Server handle objects tables.</param>
        private void ParseResponseBuffer(byte[] rgbOut, out RPC_HEADER_EXT[] rpcHeaderExts, out byte[][] rops, out uint[][] serverHandleObjectsTables)
        {
            List<RPC_HEADER_EXT> rpcHeaderExtList = new List<RPC_HEADER_EXT>();
            List<byte[]> ropList = new List<byte[]>();
            List<uint[]> serverHandleObjectList = new List<uint[]>();
            IntPtr ptr = IntPtr.Zero;

            int index = 0;
            bool isEnd = false;
            do
            {
                // Parse RPC header ext.
                RPC_HEADER_EXT rpcHeaderExt;
                ptr = Marshal.AllocHGlobal(AdapterHelper.RPCHeaderExtlength);
                                                                                                                                                                             
                // Release ptr in final sub-statement to make sure the resources will be released even if an exception occurs
                try
                {
                    Marshal.Copy(rgbOut, index, ptr, AdapterHelper.RPCHeaderExtlength);
                    rpcHeaderExt = (RPC_HEADER_EXT)Marshal.PtrToStructure(ptr, typeof(RPC_HEADER_EXT));
                    isEnd = (rpcHeaderExt.Flags & (ushort)RpcHeaderExtFlags.Last) == (ushort)RpcHeaderExtFlags.Last;
                    rpcHeaderExtList.Add(rpcHeaderExt);
                    index += AdapterHelper.RPCHeaderExtlength;
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }

                // Parse payload
                // Parse ropSize
                ushort ropSize = BitConverter.ToUInt16(rgbOut, index);
                index += ConstValues.RopSizeInRopInputOutputBufferSize;

                if ((ropSize - ConstValues.RopSizeInRopInputOutputBufferSize) > 0)
                {
                    // Parse ROP
                    byte[] rop = new byte[ropSize - ConstValues.RopSizeInRopInputOutputBufferSize];
                    Array.Copy(rgbOut, index, rop, 0, ropSize - ConstValues.RopSizeInRopInputOutputBufferSize);
                    ropList.Add(rop);
                    index += ropSize - ConstValues.RopSizeInRopInputOutputBufferSize;
                }

                // Parse server handle objects table
                // Each server handle object is 32 bytes
                Site.Assert.IsTrue(
                    (rpcHeaderExt.Size - ropSize) % sizeof(uint) == 0, "Server object handle should be uint32 array. The actual size of the server object handle is {0}.", rpcHeaderExt.Size - ropSize);

                int count = (rpcHeaderExt.Size - ropSize) / sizeof(uint);
                if (count > 0)
                {
                    uint[] sohs = new uint[count];
                    for (int counter = 0; counter < count; counter++)
                    {
                        sohs[counter] = BitConverter.ToUInt32(rgbOut, index);
                        index += sizeof(uint);
                    }
                                                                                                                                                                                                
                    serverHandleObjectList.Add(sohs);
                }
                                                                                                                                                                             
                // End parse payload
            } 
            while (!isEnd);

            rpcHeaderExts = rpcHeaderExtList.ToArray();
            rops = ropList.Count > 0 ? ropList.ToArray() : null;
            serverHandleObjectsTables = serverHandleObjectList.ToArray();
        }

        /// <summary>
        /// Registers ROPs' deserializer.
        /// </summary>
        private void RegisterROPDeserializer()
        {
            OxcropsClient ropsClient = new OxcropsClient();
            ropsClient.RegisterROPDeserializer();
        }

        /// <summary>
        /// Parses the rgbAuxOut field for EcDoConnect and EcDoRpcExt2.
        /// </summary>
        /// <param name="rgbAuxOut">The rgbAuxOut field returned by EcDoConnect or EcDoRpcExt2.</param>
        /// <returns>Array of AUX structures that can be returned from server.</returns>
        private List<AUX_SERVER_TOPOLOGY_STRUCTURE> ParseRgbAuxOut(byte[] rgbAuxOut)
        {
            List<AUX_SERVER_TOPOLOGY_STRUCTURE> auxOutList = new List<AUX_SERVER_TOPOLOGY_STRUCTURE>();
            AUX_SERVER_TOPOLOGY_STRUCTURE auxOut = new AUX_SERVER_TOPOLOGY_STRUCTURE();

            // Skip the RPC_HEADER_EXT
            int count = AdapterHelper.RPCHeaderExtlength;
            AUX_HEADER aux_header;

            do
            {
                // Parse one AUX structures
                aux_header.Size = BitConverter.ToUInt16(rgbAuxOut, count);
                aux_header.Version = rgbAuxOut[count + ConstValues.AuxHeaderSizeByteSize];
                aux_header.Type = rgbAuxOut[count + ConstValues.AuxHeaderSizeByteSize + ConstValues.AuxHeaderVersionByteSize];
                auxOut.Header = aux_header;
                auxOut.Payload = new byte[aux_header.Size - ConstValues.AuxHeaderSize];
                Array.Copy(rgbAuxOut, count + ConstValues.AuxHeaderSize, auxOut.Payload, 0, aux_header.Size - ConstValues.AuxHeaderSize);
                auxOutList.Add(auxOut);
                count += aux_header.Size;
            } 
            while (count < BitConverter.ToInt16(rgbAuxOut, ConstValues.AuxHeaderSize));

            return auxOutList;
        }
        #endregion
    }
}