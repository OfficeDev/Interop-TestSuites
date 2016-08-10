namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXCROPS protocol adapter.
    /// </summary>
    public partial class MS_OXCROPSAdapter : ManagedAdapterBase, IMS_OXCROPSAdapter
    {
        #region Public Fields for consts definitions.

        /// <summary>
        /// Definition for several points in MS-OXCROPS referring 0xBABE.
        /// </summary>
        public const uint BufferSize = 0xBABE;

        /// <summary>
        /// This is used to set Flags of RPC_HEADER_EXT,which indicates it is compressed.
        /// </summary>
        public const ushort CompressedForFlagsOfHeader = 0x0001;

        /// <summary>
        /// Definition for default value of Output handle. 
        /// </summary>
        public const uint DefaultOutputHandle = 0xFFFFFFFF;

        /// <summary>
        /// FolderId (8 bytes): 64-bit identifier. This field MUST be set to 0x0000000000000000
        /// </summary>
        public const ulong FolderIdForRopSynchronizationImportHierarchyChange = 0x0000000000000000;

        /// <summary>
        /// Server object handle value 0xFFFFFFFF is used to initialize unused entries of a Server object handle table.
        /// </summary>
        public const uint HandleValueForUnusedEntries = 0xFFFFFFFF;

        /// <summary>
        /// Definition for invalid input handle. 
        /// </summary>
        public const uint InvalidInputHandle = 0xFFFFFFFF;

        /// <summary>
        /// This is used to set Flags of RPC_HEADER_EXT,which indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT.
        /// </summary>
        public const ushort LastForFlagsOfHeader = 0x0004;

        /// <summary>
        /// The maximum value of pcbOut 
        /// </summary>
        public const uint MaxPcbOut = 0x40000;

        /// <summary>
        /// Definition for MaxrgbOut,which specifies the maximum size of the rgbOut buffer to place in Response
        /// </summary>
        public const int MaxRgbOut = 0x10008;

        /// <summary>
        ///  MessageId (8 bytes): This field MUST be set to 0x0000000000000000.
        /// </summary>
        public const ulong MessageIdForRops = 0x0000000000000000;

        /// <summary>
        /// Definition for the Null-Terminating string.
        /// </summary>
        public const char NullTerminateCharacter = '\0';

        /// <summary>
        /// Definition for PayloadLen which indicates the length of the field that represents the length of payload.
        /// </summary>
        public const int PayloadLen = 0x2;

        /// <summary>
        /// This is used to set the max size of the rgbAuxOut.
        /// </summary>
        public const uint PcbAuxOut = 0x1008;

        /// <summary>
        /// This flags indicates client requests server to not compress or XOR payload of rgbOut and rgbAuxOut.
        /// </summary>
        public const uint PulFlags = 0x00000003;

        /// <summary>
        /// Unsigned 64-bit integer. This value specifies the number of bytes read from the source object or written to the destination object.
        /// </summary>
        public const ulong ReadOrWrittenByteCountForRopCopyToStream = 0x0000000000000000;

        /// <summary>
        /// Definition the one reserved byte: 0x00
        /// </summary>
        public const byte ReservedOneByte = 0x00;

        /// <summary>
        /// Definition the two reserved bytes: 0x0000
        /// </summary>
        public const ushort ReservedTwoBytes = 0x0000;

        /// <summary>
        /// Definition for ReturnValue of PulFlags,which MUST be set to 0x00000000.
        /// </summary>
        public const int ReturnValueForPulFlags = 0x00000000;

        /// <summary>
        /// Definition for the possible return value of RopFastTransferSourceGetBufferResponse: 0x00000480.
        /// </summary>
        public const uint ReturnValueForRopFastTransferSourceGetBufferResponse = 0x00000480;

        /// <summary>
        /// Definition for the possible return value of RopMoveFolderResponse and RopMoveCopyMessagesResponse: 0x00000503.
        /// </summary>
        public const uint ReturnValueForRopMoveFolderResponseAndMoveCopyMessage = 0x00000503;

        /// <summary>
        /// Definition for ReturnValue of RopQueryNamedProperties: 0x00040380.
        /// </summary>
        public const uint ReturnValueForRopQueryNamedProperties = 0x00040380;

        /// <summary>
        /// Definition for the possible return value of Redirect response: 0x00000478.
        /// </summary>
        public const uint WrongServer = 0x00000478;

        /// <summary>
        /// Definition for ReturnValue of ret, which MUST be set to 0x0.
        /// </summary>
        public const uint ReturnValueForRet = 0x0;

        /// <summary>
        /// Definition for ReturnValue of success response: 0x00000000.
        /// </summary>
        public const uint SuccessReturnValue = 0x00000000;

        /// <summary>
        /// Definition for RopSize which specifies the size of both this field and the RopsList field.
        /// </summary>
        public const ushort RopSize = 0x2;

        /// <summary>
        /// Definition for Version of RpcHeaderExt,this value MUST be set to 0x00.
        /// </summary>
        public const ushort VersionOfRpcHeaderExt = 0x00;

        /// <summary>
        /// This is used to set Flags of RPC_HEADER_EXT,which indicates it is obfuscated.
        /// </summary>
        public const ushort XorMagicForFlagsOfHeader = 0x0002;

        #endregion

        #region Private Fields

        /// <summary>
        /// Length of RPC_HEADER_EXT
        /// </summary>
        private static readonly int RPCHEADEREXTLEN = Marshal.SizeOf(typeof(RPC_HEADER_EXT));

        /// <summary>
        /// Whether import the common configuration file.
        /// </summary>
        private static bool commonConfigImported = false;

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// Status of connection.
        /// </summary>
        private bool isConnected;

        #endregion

        #region IMS_OXCROPSAdapter members

        /// <summary>
        /// Connect to the server for RPC calling.
        /// This method is defined as a direct way to connect to server with specific parameters.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">the type of connection</param>
        /// <param name="userDN">User DN used to connect server</param>
        /// <param name="domain">Domain name</param>
        /// <param name="userName">User name used to logon</param>
        /// <param name="password">User Password</param>
        /// <returns>Result of connecting.</returns>
        public bool RpcConnect(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password)
        {
            bool ret = this.oxcropsClient.Connect(
                server,
                connectionType,
                userDN,
                domain,
                userName,
                password);
            this.isConnected = ret;
            return ret;
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>Result of disconnecting.</returns>
        public bool RpcDisconnect()
        {
            if (this.isConnected)
            {
                // Disconnect the RPC link to server.
                bool ret = this.oxcropsClient.Disconnect();
                if (ret)
                {
                    this.isConnected = false;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Set auto redirect value in RPC context
        /// If setting this to true, the RPC server will return EcWrongServer error (0x478). And the request will be redirected to designated server.
        /// If setting this to false, the RPC server will return EcWrongServer error (0x478). But the request will not be redirected.
        /// </summary>
        /// <param name="option">true indicates enable auto redirect, false indicates disable auto redirect</param>
        public void SetAutoRedirect(bool option)
        {
            this.oxcropsClient.MapiContext.AutoRedirect = option;
        }

        /// <summary>
        /// Method which executes single ROP.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <param name="returnValue">The return value of the ROP method.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> ProcessSingleRopWithReturnValue(
            ISerializable ropRequest,
            uint inputObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType,
            out uint returnValue)
        {
            List<ISerializable> requestRops = null;
            if (ropRequest != null)
            {
                requestRops = new List<ISerializable>
                {
                    ropRequest
                };
            }

            List<uint> requestSOH = new List<uint>
            {
                inputObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle and set default value to 0xFFFFFFFF.
                requestSOH.Add(DefaultOutputHandle);
            }

            if (this.IsInvalidInputHandleNeeded(ropRequest, expectedRopResponseType))
            {
                // Add an invalid input handle in request and set its value to 0xFFFFFFFF.
                requestSOH.Add(InvalidInputHandle);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, MaxRgbOut);
            returnValue = ret;
            if (ret != OxcRpcErrorCode.ECNone && ret != 1726)
            {
                Site.Assert.AreEqual<RopResponseType>(RopResponseType.RPCError, expectedRopResponseType, "Unexpected RPC error {0} occurred.", ret);
                return responseSOHs;
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

            if (ropRequest.GetType() == typeof(RopReleaseRequest))
            {
                return responseSOHs;
            }

            if (response.GetType() == typeof(RopSaveChangesMessageResponse) && ((RopSaveChangesMessageResponse)response).ReturnValue == 0x80040401)
            {
                return responseSOHs;
            }

            if (response != null)
            {
                try
                {
                    this.VerifyAdapterCaptureCode(expectedRopResponseType, response, ropRequest);
                }
                catch (TargetInvocationException invocationEx)
                {
                    Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                    if (invocationEx.InnerException != null)
                    {
                        throw invocationEx.InnerException;
                    }
                }
                catch (NullReferenceException nullEx)
                {
                    Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
                }
            }

            return responseSOHs;
        }

        /// <summary>
        /// Method which executes single ROP.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> ProcessSingleRop(
            ISerializable ropRequest,
            uint inputObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType)
        {
            uint returnValue;
            List<List<uint>> responseSOHs = this.ProcessSingleRopWithReturnValue(ropRequest, inputObjHandle, ref response, ref rawData, expectedRopResponseType, out returnValue);
            if (returnValue == 1726)
            {
                Site.Assert.AreEqual<RopResponseType>(RopResponseType.RPCError, expectedRopResponseType, "Unexpected RPC error {0} occurred.", returnValue);
            }

            return responseSOHs;
        }

        /// <summary>
        /// Method which executes single ROP operation with the maximum size of the rgbOut buffer set as pcbOut.
        /// For more detail about rgbOut and pcbOut, see [MS-OXCRPC].
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <param name="pcbOut">The maximum size of the rgbOut buffer to place Response in.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> ProcessSingleRopWithOptionResponseBufferSize(
            ISerializable ropRequest,
            uint inputObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType,
            uint pcbOut)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>
            {
                inputObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle and set default value to 0xFFFFFFFF.
                requestSOH.Add(DefaultOutputHandle);
            }

            if (this.IsInvalidInputHandleNeeded(ropRequest, expectedRopResponseType))
            {
                // Add an invalid input handle in request and set its value to 0xFFFFFFFF.
                requestSOH.Add(InvalidInputHandle);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, pcbOut);
            if (ret != OxcRpcErrorCode.ECNone)
            {
                Site.Assert.AreEqual<RopResponseType>(RopResponseType.RPCError, expectedRopResponseType, "Unexpected RPC error {0} occurred.", ret);
                return null;
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

            if (ropRequest.GetType() == typeof(RopReleaseRequest))
            {
                return responseSOHs;
            }

            try
            {
                this.VerifyAdapterCaptureCode(expectedRopResponseType, response, ropRequest);
            }
            catch (TargetInvocationException invocationEx)
            {
                Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                if (invocationEx.InnerException != null)
                {
                    throw invocationEx.InnerException;
                }
            }
            catch (NullReferenceException nullEx)
            {
                Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
            }

            return responseSOHs;
        }

        /// <summary>
        /// Method which executes multiple ROPs.
        /// </summary>
        /// <param name="requestRops">ROP request objects.</param>
        /// <param name="inputObjHandles">Server object handles in request.</param>
        /// <param name="responseRops">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">The expected response type.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> ProcessMutipleRops(
            List<ISerializable> requestRops,
            List<uint> inputObjHandles,
            ref List<IDeserializable> responseRops,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType)
        {
            List<uint> requestSOH = new List<uint>();
            for (int i = 0; i < inputObjHandles.Count; i++)
            {
                requestSOH.Add(inputObjHandles[i]);
            }

            List<List<uint>> responseSOHs = new List<List<uint>>();
            uint ret = this.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, MaxRgbOut);
            if (ret != OxcRpcErrorCode.ECNone)
            {
                Site.Assert.AreEqual<RopResponseType>(RopResponseType.RPCError, expectedRopResponseType, "Unexpected RPC error {0} occurred.", ret);
                return responseSOHs;
            }

            int numOfReqs = requestRops.Count;
            int numOfRess = responseRops.Count;

            for (int reqIndex = 0, resIndex = 0; reqIndex < numOfReqs && resIndex < numOfRess; resIndex++, reqIndex++)
            {
                while (requestRops[reqIndex].GetType() == typeof(RopReleaseRequest) && (reqIndex < numOfReqs - 1))
                {
                    reqIndex++;
                }

                try
                {
                    Type reqType = requestRops[reqIndex].GetType();
                    string resName = responseRops[resIndex].GetType().Name;

                    // The word "Response" takes 8 length.
                    string ropName = resName.Substring(0, resName.Length - 8);
                    Type adapterType = typeof(MS_OXCROPSAdapter);

                    // Call capture code using reflection mechanism
                    // The code followed is to construct the verify method name of capture code and then call this method through reflection.
                    MethodInfo method = null;
                    string verifyMethodName = "Verify" + ropName + "SuccessResponse";
                    method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                    if (method == null)
                    {
                        verifyMethodName = "Verify" + ropName + "FailureResponse";
                        method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                    }

                    if (method == null)
                    {
                        verifyMethodName = "Verify" + ropName + "Response";
                        method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                    }

                    if (method != null)
                    {
                        ParameterInfo[] paraInfos = method.GetParameters();
                        int paraNum = paraInfos.Length;
                        object[] paraObjects = new object[paraNum];
                        paraObjects[0] = responseRops[resIndex];
                        for (int i = 1; i < paraNum; i++)
                        {
                            FieldInfo fieldInReq = reqType.GetField(
                                paraInfos[i].Name,
                                BindingFlags.IgnoreCase
                                | BindingFlags.DeclaredOnly
                                | BindingFlags.Public
                                | BindingFlags.NonPublic
                                | BindingFlags.GetField
                                | BindingFlags.Instance);
                            if (fieldInReq == null)
                            {
                                foreach (ISerializable req in requestRops)
                                {
                                    Type type = req.GetType();
                                    fieldInReq = type.GetField(
                                        paraInfos[i].Name,
                                        BindingFlags.IgnoreCase
                                        | BindingFlags.DeclaredOnly
                                        | BindingFlags.Public
                                        | BindingFlags.NonPublic
                                        | BindingFlags.GetField
                                        | BindingFlags.Instance);
                                    if (fieldInReq != null)
                                    {
                                        paraObjects[i] = fieldInReq.GetValue(req);
                                    }
                                }
                            }
                            else
                            {
                                paraObjects[i] = fieldInReq.GetValue(requestRops[reqIndex]);
                            }
                        }

                        method.Invoke(this, paraObjects);
                    }
                }
                catch (TargetInvocationException invocationEx)
                {
                    Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                    if (invocationEx.InnerException != null)
                    {
                        throw invocationEx.InnerException;
                    }
                }
                catch (NullReferenceException nullEx)
                {
                    Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
                }

                if (resIndex < numOfRess - 1)
                {
                    if (responseRops[resIndex + 1].GetType() == typeof(RopNotifyResponse) || responseRops[resIndex + 1].GetType() == typeof(RopPendingResponse))
                    {
                        reqIndex--;
                    }
                }
            }

            return responseSOHs;
        }

        /// <summary>
        /// Method which executes single ROP with multiple server objects.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="inputObjHandles">Server object handles in request.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> ProcessSingleRopWithMutipleServerObjects(
            ISerializable ropRequest,
            List<uint> inputObjHandles,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>();
            for (int i = 0; i < inputObjHandles.Count; i++)
            {
                requestSOH.Add(inputObjHandles[i]);
            }

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                 // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(DefaultOutputHandle);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            uint ret = this.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, MaxRgbOut);
            if (ret != OxcRpcErrorCode.ECNone)
            {
                Site.Assert.AreEqual<RopResponseType>(RopResponseType.RPCError, expectedRopResponseType, "Unexpected RPC error {0} occurred.", ret);
                return responseSOHs;
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

            if (ropRequest.GetType() == typeof(RopReleaseRequest))
            {
                return responseSOHs;
            }

            try
            {
                this.VerifyAdapterCaptureCode(expectedRopResponseType, response, ropRequest);
            }
            catch (TargetInvocationException invocationEx)
            {
                Site.Log.Add(LogEntryKind.Debug, invocationEx.Message);
                if (invocationEx.InnerException != null)
                {
                    throw invocationEx.InnerException;
                }
            }
            catch (NullReferenceException nullEx)
            {
                Site.Log.Add(LogEntryKind.Debug, nullEx.Message);
            }

            return responseSOHs;
        }

        #endregion

        #region IAdapter members
        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCROPS";
            if (!commonConfigImported)
            {
                Common.MergeConfiguration(this.Site);
                commonConfigImported = true;
            }

            // Initialize OxcropsClient instance.
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        #endregion

        #region Private Function

        #region Functions for Adapter

        /// <summary>
        /// Check whether the array of byte is null terminated ASCII string.
        /// </summary>
        /// <param name="buffer">The array of byte which to be checked whether  null terminated ASCII string</param>
        /// <returns>A boolean type indicating whether the passed string is a null-terminated ASCII string.</returns>
        private bool IsNullTerminatedASCIIStr(byte[] buffer)
        {
            int len = buffer.Length;

            // Check null terminate.
            bool isNullTerminated = buffer[len - 1] == 0x00;
            bool isASCIIString = true;
            for (int i = 0; i < buffer.Length; i++)
            {
                // ASCII between 0x00 and 0x7F
                if (buffer[i] >= 0x00 && buffer[i] <= 0x7F)
                {
                    continue;
                }
                else
                {
                    isASCIIString = false;
                    break;
                }
            }

            return isNullTerminated && isASCIIString;
        }

        /// <summary>
        /// Verify whether the GUID bytes is GUID or not
        /// </summary>
        /// <param name="guidBytes">An array of bytes.</param>
        /// <returns>If it is GUID, return true, else return false.</returns>
        private bool IsGUID(byte[] guidBytes)
        {
            bool isGUID = false;

            // Check GUID length with 16.
            if (guidBytes.Length == 16)
            {
                Guid guid = new Guid(guidBytes);

                // GUID format check regExpression.
                string regexPatten = @"^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\"
                    + @"-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$";
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(regexPatten);
                string guidStr = guid.ToString();
                isGUID = regex.IsMatch(guidStr);
            }

            return isGUID;
        }

        /// <summary>
        /// Verify adapter capture code using reflection mechanism.
        /// </summary>
        /// <param name="expectedRopResponseType">The Expected ROP response type.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="ropRequest">ROP request object.</param>
        private void VerifyAdapterCaptureCode(RopResponseType expectedRopResponseType, IDeserializable response, ISerializable ropRequest)
        {
            string resName = response.GetType().Name;

            // The word "Response" takes 8 length.
            string ropName = resName.Substring(0, resName.Length - 8);
            Type adapterType = typeof(MS_OXCROPSAdapter);

            // Call capture code using reflection mechanism
            // The code followed is to construct the verify method name of capture code and then call this method through reflection.
            string verifyMethodName = string.Empty;
            if (expectedRopResponseType == RopResponseType.SuccessResponse)
            {
                verifyMethodName = "Verify" + ropName + "SuccessResponse";
            }
            else if (expectedRopResponseType == RopResponseType.FailureResponse)
            {
                verifyMethodName = "Verify" + ropName + "FailureResponse";
            }
            else if (expectedRopResponseType == RopResponseType.Response)
            {
                verifyMethodName = "Verify" + ropName + "Response";
            }
            else if (expectedRopResponseType == RopResponseType.NullDestinationFailureResponse)
            {
                verifyMethodName = "Verify" + ropName + "NullDestinationFailureResponse";
            }
            else if (expectedRopResponseType == RopResponseType.RedirectResponse)
            {
                verifyMethodName = "Verify" + ropName + "RedirectResponse";
            }

            Type reqType = ropRequest.GetType();
            MethodInfo method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
            if (method == null)
            {
                if (expectedRopResponseType == RopResponseType.SuccessResponse || expectedRopResponseType == RopResponseType.FailureResponse)
                {
                    verifyMethodName = "Verify" + ropName + "Response";
                    method = adapterType.GetMethod(verifyMethodName, BindingFlags.NonPublic | BindingFlags.Instance);
                }
            }

            if (method != null)
            {
                ParameterInfo[] paraInfos = method.GetParameters();
                int paraNum = paraInfos.Length;
                object[] paraObjects = new object[paraNum];
                paraObjects[0] = response;
                for (int i = 1; i < paraNum; i++)
                {
                    FieldInfo fieldInReq = reqType.GetField(
                        paraInfos[i].Name,
                        BindingFlags.IgnoreCase
                        | BindingFlags.DeclaredOnly
                        | BindingFlags.Public
                        | BindingFlags.NonPublic
                        | BindingFlags.GetField
                        | BindingFlags.Instance);
                    paraObjects[i] = fieldInReq.GetValue(ropRequest);
                }

                method.Invoke(this, paraObjects);
            }
        }

        /// <summary>
        /// Check whether the default invalid handle is needed.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="expectedRopResponseType">The expected ROP response type.</param>
        /// <returns>Return true if the default handle is needed in the request, otherwise return false.</returns>
        private bool IsInvalidInputHandleNeeded(ISerializable ropRequest, RopResponseType expectedRopResponseType)
        {
            if (!Common.IsOutputHandleInRopRequest(ropRequest) && expectedRopResponseType == RopResponseType.FailureResponse)
            {
                byte[] request = ropRequest.Serialize();

                // The default handle is also needed by some cases to verify the failure response caused by an invalid input handle.
                // The input handle index is the third byte and its value is 1 in this test suite for this situation.
                byte inputHandleIndex = request[2];
                if (inputHandleIndex == 1)
                {
                    return true;
                }
            }

            return false;
        }

        #endregion

        #region Functions for Process
        /// <summary>
        /// The method parses response buffer.
        /// </summary>
        /// <param name="rgbOut">The ROP response payload.</param>
        /// <param name="rpcHeaderExts">RPC header extension.</param>
        /// <param name="rops">ROPs in response.</param>
        /// <param name="serverHandleObjectsTables">Server handle objects tables</param>
        private void ParseResponseBuffer(byte[] rgbOut, out RPC_HEADER_EXT[] rpcHeaderExts, out byte[][] rops, out uint[][] serverHandleObjectsTables)
        {
            List<RPC_HEADER_EXT> rpcHeaderExtList = new List<RPC_HEADER_EXT>();
            List<byte[]> ropList = new List<byte[]>();
            List<uint[]> serverHandleObjectList = new List<uint[]>();
            IntPtr ptr = IntPtr.Zero;

            int index = 0;
            bool end = false;
            do
            {
                // Parse RPC header extension
                RPC_HEADER_EXT rpcHeaderExt;
                ptr = Marshal.AllocHGlobal(RPCHEADEREXTLEN);
                try
                {
                    Marshal.Copy(rgbOut, index, ptr, RPCHEADEREXTLEN);
                    rpcHeaderExt = (RPC_HEADER_EXT)Marshal.PtrToStructure(ptr, typeof(RPC_HEADER_EXT));

                    end = (rpcHeaderExt.Flags & LastForFlagsOfHeader) == LastForFlagsOfHeader;
                    rpcHeaderExtList.Add(rpcHeaderExt);
                    index += RPCHEADEREXTLEN;
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }

                #region  Start parse payload
                // Parse ropSize
                ushort ropSize = BitConverter.ToUInt16(rgbOut, index);
                index += sizeof(ushort);

                if ((ropSize - sizeof(ushort)) > 0)
                {
                    // Parse ROP
                    byte[] rop = new byte[ropSize - sizeof(ushort)];
                    Array.Copy(rgbOut, index, rop, 0, ropSize - sizeof(ushort));
                    ropList.Add(rop);
                    index += ropSize - sizeof(ushort);
                }

                // Parse server handle objects table
                Site.Assert.IsTrue(
                    (rpcHeaderExt.Size - ropSize) % sizeof(uint) == 0, "server object handle should be uint32 array");

                int count = (rpcHeaderExt.Size - ropSize) / sizeof(uint);
                if (count > 0)
                {
                    uint[] sohs = new uint[count];
                    for (int k = 0; k < count; k++)
                    {
                        sohs[k] = BitConverter.ToUInt32(rgbOut, index);
                        index += sizeof(uint);
                    }

                    serverHandleObjectList.Add(sohs);
                }
                #endregion
            }
            while (!end);

            rpcHeaderExts = rpcHeaderExtList.ToArray();
            rops = ropList.ToArray();
            serverHandleObjectsTables = serverHandleObjectList.ToArray();
        }

        /// <summary>
        /// The method creates ROPs request.
        /// </summary>
        /// <param name="requestROPs">ROPs in request.</param>
        /// <param name="requestSOHTable">Server object handles table.</param>
        /// <returns>The ROPs request.</returns>
        private byte[] BuildRequestBuffer(List<ISerializable> requestROPs, List<uint> requestSOHTable)
        {
            // Definition for PayloadLen which indicates the length of the field that represents the length of payload.
            int payloadLen = PayloadLen;
            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    payloadLen += requestROP.Size();
                }
            }

            Site.Assert.IsTrue(
                payloadLen < ushort.MaxValue,
                "The number of bytes in this field MUST be 2 bytes less than the value specified in the RopSize field.");

            ushort ropSize = (ushort)payloadLen;

            if (requestSOHTable != null)
            {
                payloadLen += requestSOHTable.Count * sizeof(uint);
            }

            byte[] requestBuffer = new byte[RPCHEADEREXTLEN + payloadLen];
            int index = 0;

            // Construct RPC header extension buffer
            RPC_HEADER_EXT rpcHeaderExt = new RPC_HEADER_EXT
            {
                Version = VersionOfRpcHeaderExt, // There is only one version of the header at this time so this value MUST be set to 0x00.
                Flags = LastForFlagsOfHeader, // Last (0x04) indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT. 
                Size = (ushort)payloadLen
            };
            rpcHeaderExt.SizeActual = rpcHeaderExt.Size;

            IntPtr ptr = Marshal.AllocHGlobal(RPCHEADEREXTLEN);
            try
            {
                Marshal.StructureToPtr(rpcHeaderExt, ptr, true);
                Marshal.Copy(ptr, requestBuffer, index, RPCHEADEREXTLEN);
                index += RPCHEADEREXTLEN;
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }

            // RopSize's type is ushort. So the offset will be 2.
            Array.Copy(BitConverter.GetBytes(ropSize), 0, requestBuffer, index, 2);
            index += 2;

            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    Array.Copy(requestROP.Serialize(), 0, requestBuffer, index, requestROP.Size());
                    index += requestROP.Size();
                }
            }

            if (requestSOHTable != null)
            {
                foreach (uint serverHandle in requestSOHTable)
                {
                    Array.Copy(BitConverter.GetBytes(serverHandle), 0, requestBuffer, index, sizeof(uint));
                    index += sizeof(uint);
                }
            }

            return requestBuffer;
        }

        /// <summary>
        /// Do EcDoRPCExt2 call.
        /// </summary>
        /// <param name="requestROPs">ROP request objects.</param>
        /// <param name="requestSOHTable">ROP request server object handle table.</param>
        /// <param name="responseROPs">ROP response objects.</param>
        /// <param name="responseSOHTable">ROP response server object handle table.</param>
        /// <param name="rawData">The response payload bytes.</param>
        /// <param name="pcbOut">The maximum size of the rgbOut buffer to place Response in.</param>
        /// <returns>0 indicates success, other values indicate failure. </returns>
        private uint RopCall(
            List<ISerializable> requestROPs,
            List<uint> requestSOHTable,
            ref List<IDeserializable> responseROPs,
            ref List<List<uint>> responseSOHTable,
            ref byte[] rawData,
            uint pcbOut)
        {
            byte[] rgbIn = this.BuildRequestBuffer(requestROPs, requestSOHTable);
            uint ret = this.oxcropsClient.RopCall(requestROPs, requestSOHTable, ref responseROPs, ref responseSOHTable, ref rawData, pcbOut);

            if (ret != OxcRpcErrorCode.ECNone)
            {
                // Verify error for RPC
                if (this.oxcropsClient.IsReservedRopId(rgbIn[10]))
                {
                    this.VerifyRPCErrorEncounterReservedRopIds(rgbIn[10], (uint)ret);
                    return ret;
                }
                else
                {
                    // RopOpenFolder
                    if (rgbIn[10] == Convert.ToByte(RopId.RopOpenFolder))
                    {
                        if (ret == OxcRpcErrorCode.ECRpcFormat)
                        {
                            this.VerifyRPCErrorEncounterUnableParseRequest((uint)ret);
                        }

                        return ret;
                    }

                    // RopReadStream
                    if (rgbIn[10] == Convert.ToByte(RopId.RopReadStream))
                    {
                        if (ret == OxcRpcErrorCode.ECRpcFormat)
                        {
                            this.VerifyMaximumByteCountExceedError((uint)ret);
                        }

                        return ret;
                    }
                }
            }

            if (ret == OxcRpcErrorCode.ECRpcFormat)
            {
                Site.Assert.Fail("Error RPC Format");
            }

            if (ret == OxcRpcErrorCode.ECResponseTooBig)
            {
                this.VerifyFailRPCForMaxPcbOut(ret);
                this.VerifyFailRPCForInsufficientOutputBuffer(ret);
                return ret;
            }

            Site.Assert.AreEqual<uint>(ReturnValueForRet, ret, "If the response is success, the return value is 0x0.");

            this.VerifyTransport();

            RPC_HEADER_EXT[] rpcHeaderExts;
            byte[][] rops;
            uint[][] serverHandleObjectsTables;
            this.ParseResponseBuffer(rawData, out rpcHeaderExts, out rops, out serverHandleObjectsTables);
            Site.Assert.AreEqual<int>(1, rpcHeaderExts.Length, "Support one rpc/payload only");

            if (responseSOHTable.Count > 0)
            {
                ushort ropSize = BitConverter.ToUInt16(rawData, RPCHEADEREXTLEN);

                // Verify each ROP response buffer structure 
                this.VerifyMessageSyntaxRequestAndResponseBuffer(ropSize, rops, responseSOHTable[0], rawData);

                // Verify Message Processing Events and Sequencing Rules
                this.VerifyMessageProcessingEventsAndSequencingRules(responseSOHTable[0]);
            }

            return ret;
        }

        #endregion
        #endregion
    }
}