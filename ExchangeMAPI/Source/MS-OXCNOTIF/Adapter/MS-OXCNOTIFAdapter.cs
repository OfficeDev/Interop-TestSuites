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
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Net;
    using System.Net.Sockets;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXCNOTIF Protocol Adapter.
    /// </summary>
    public partial class MS_OXCNOTIFAdapter : ManagedAdapterBase, IMS_OXCNOTIFAdapter
    {
        /// <summary>
        /// Whether import the common configuration file.
        /// </summary>
        private static bool commonConfigImported;

        /// <summary>
        /// The instance of OxcropsClient.
        /// </summary>
        private OxcropsClient oxcRopsClient;

        /// <summary>
        /// The instance of OxcropsClient, used to trigger the notifications.
        /// </summary>
        private OxcropsClient oxcRopsTrigger;

        /// <summary>
        /// The instance of OxcropsClient, used to register the notifications.
        /// </summary>
        private OxcropsClient oxcRopsRegister;

        /// <summary>
        /// Store the server object handles returned by response.
        /// </summary>
        private List<List<uint>> responseSOHs = new List<List<uint>>();

        /// <summary>
        /// The flags out pointer.
        /// </summary>
        private IntPtr pulFlagsOut;

        /// <summary>
        /// The UdpClient instance used to receive push notification.
        /// </summary>
        private UdpClient udp;

        /// <summary>
        /// Gets or sets Logon Handle.
        /// </summary>
        public uint LogonHandle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the session is connect.
        /// </summary>
        public bool IsConnected
        {
            get
            {
                return this.oxcRopsClient.IsConnected;
            }

            set
            {
                this.oxcRopsClient.IsConnected = value;
            }
        }

        /// <summary>
        /// Gets or sets RPC Context.
        /// </summary>
        public IntPtr RPCContext
        {
            get
            {
                return this.oxcRopsClient.CXH;
            }

            set
            {
                this.oxcRopsClient.CXH = value;
            }
        }

        /// <summary>
        /// Gets or sets the logon on which the operation is performed. Default value is 0.
        /// </summary>
        private byte LogonId { get; set; }

        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXCNOTIF";

            if (!commonConfigImported)
            {
                Common.MergeConfiguration(this.Site);
                commonConfigImported = true;
            }

            this.oxcRopsRegister = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
            this.oxcRopsTrigger = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
            this.oxcRopsClient = this.oxcRopsTrigger;
        }

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection.</param>
        /// <returns>If the behavior of connecting server is successful, the server will return true; otherwise, return false.</returns>
        public bool DoConnect(ConnectionType connectionType)
        {
            return this.oxcRopsClient.Connect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    connectionType,
                    Common.GetConfigurationPropertyValue("User1Essdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("User1Name", this.Site),
                    Common.GetConfigurationPropertyValue("User1Password", this.Site));
        }

        /// <summary>
        /// Disconnect the connection with Server.
        /// </summary>
        /// <returns>The return value indicates the disconnection state.</returns>
        public bool DoDisconnect()
        {
            return this.oxcRopsClient.Disconnect();
        }

        /// <summary>
        /// Switch the instance of OxcropsClient.
        /// </summary>
        public void SwitchSessionContext()
        {
            if (this.oxcRopsClient == this.oxcRopsTrigger)
            {
                this.oxcRopsClient = this.oxcRopsRegister;
            }
            else
            {
                this.oxcRopsClient = this.oxcRopsTrigger;
            }
        }

        /// <summary>
        /// Register a callback address on the server which will be used to push notification.
        /// </summary>
        /// <param name="addressFamily">The AddressFamily type specifies which IP family to use.</param>
        /// <param name="port">The UDP port that will receive the push notification.</param>
        /// <param name="opaque">The opaque client-generated context data that is sent back to the client at the callback address.</param>
        /// <returns>If the method succeeds, the return value is 0. If the method fails, the return value is an implementation-specific error code.</returns>
        public uint EcRRegisterPushNotification(AddressFamily addressFamily, int port, string opaque)
        {
            IntPtr cxh = this.oxcRopsClient.CXH;
            string ipv4 = Common.GetConfigurationPropertyValue("NotificationIP", this.Site);
            string ipv6 = Common.GetConfigurationPropertyValue("NotificationIPv6", this.Site);

            byte[] rgbContext = Encoding.ASCII.GetBytes(opaque);
            ushort contextLength = (ushort)rgbContext.Length;
            uint notification;
            string ip = addressFamily == AddressFamily.AF_INET ? ipv4 : ipv6;
            Site.Assert.AreNotEqual<string>(string.Empty, ip, "The IP address should be gotten successfully.");
            uint ret = NativeMethods.EcRRegisterPushNotificationWrap(ref cxh, (short)addressFamily, ip, (ushort)port, rgbContext, contextLength, out notification);
            this.VerifyAsyncCallOnRPCTransport();
            return ret;
        }

        /// <summary>
        /// Receive push notification on the specified port.
        /// </summary>
        /// <param name="addressFamily">The AddressFamily type specifies which IP family to use.</param>
        /// <param name="port">The UDP port uses to receive the push notification.</param>
        /// <param name="opaque">The opaque data received from server.</param>
        /// <returns>True if the notification received, otherwise false.</returns>
        public bool PushNotificationReceived(AddressFamily addressFamily, int port, out string opaque)
        {
            opaque = null;
            if (this.udp == null)
            {
                try
                {
                    this.udp = new UdpClient(port, addressFamily == AddressFamily.AF_INET ? System.Net.Sockets.AddressFamily.InterNetwork : System.Net.Sockets.AddressFamily.InterNetworkV6)
                    {
                        Client =
                        {
                            ReceiveTimeout = int.Parse(Common.GetConfigurationPropertyValue("PushNotificationTimeout", this.Site))
                        }
                    };
                }
                catch (SocketException e)
                {
                    Site.Assert.Fail("The UDP port {0} fails to bind. {1}", port, e.Message);
                }
            }

            IPEndPoint remote = new IPEndPoint(IPAddress.Any, 0);
            byte[] data;
            try
            {
                data = this.udp.Receive(ref remote);
                opaque = Encoding.ASCII.GetString(data);
                this.VerifyUDPTransport();
                this.VerifyCallbackAddressForUDPDatagrams();
                return true;
            }
            catch (SocketException e)
            {
                Site.Log.Add(LogEntryKind.Debug, "The error message of receiving UDP is {0}", e.Message);
                return false;
            }
            finally
            {
                this.udp.Close();
                this.udp = null;
            }
        }

        /// <summary>
        /// Acquire an asynchronous context handle on the server which will be used in subsequent EcDoAsyncWaitEx call.
        /// </summary>
        /// <returns>The asynchronous context handle.</returns>
        public IntPtr EcDoAsyncConnectEx()
        {
            uint returnValue;
            IntPtr acxh = IntPtr.Zero;
            try
            {
                returnValue = NativeMethods.EcDoAsyncConnectEx(this.oxcRopsClient.CXH, ref acxh);
            }
            catch (SEHException ex)
            {
                returnValue = NativeMethods.RpcExceptionCode(ex);
            }

            Site.Assert.AreEqual<uint>(0, returnValue, "EcDoAsyncConnectEx should succeed");
            this.VerifyAsyncCallOnRPCTransport();
            return acxh;
        }

        /// <summary>
        /// Call EcDoAsyncWaitEx method and return immediately.
        /// </summary>
        /// <param name="acxh">The asynchronous context handle</param>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        public void BeginAsyncWait(IntPtr acxh, out IntPtr rpcAsyncHandle)
        {
            rpcAsyncHandle = NativeMethods.CreateRpcAsyncHandle();
            Site.Assert.AreNotEqual<IntPtr>(IntPtr.Zero, rpcAsyncHandle, "Get valid asynchronous handle");
            this.pulFlagsOut = Marshal.AllocHGlobal(sizeof(int));

            NativeMethods.EcDoAsyncWaitEx(rpcAsyncHandle, acxh, 0, this.pulFlagsOut);
            this.VerifyAsyncCallOnRPCTransport();
        }

        /// <summary>
        /// Get the status of EcDoAsyncWaitEx call.
        /// </summary>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        /// <returns>The status of asynchronous call</returns>
        public RPCAsyncStatus QueryAsyncWaitStatus(IntPtr rpcAsyncHandle)
        {
            RPCAsyncStatus status = (RPCAsyncStatus)NativeMethods.RpcAsyncGetCallStatus(rpcAsyncHandle);
            this.VerifyAsyncCallOnRPCTransport();
            return status;
        }

        /// <summary>
        /// Complete the EcDoAsyncWaitEx call.
        /// </summary>
        /// <param name="rpcAsyncHandle">RPC asynchronous handle</param>
        /// <param name="flagsOut">The pulFlagsOut parameter returned by EcDoAsyncWaitEx</param>
        /// <returns>The EcDoAsyncWaitEx call return value</returns>
        public int EndAsyncWait(IntPtr rpcAsyncHandle, out int flagsOut)
        {
            IntPtr reply = Marshal.AllocHGlobal(sizeof(int));
            RPCAsyncStatus status = (RPCAsyncStatus)NativeMethods.RpcAsyncCompleteCall(rpcAsyncHandle, reply);
            Site.Assert.AreEqual<RPCAsyncStatus>(RPCAsyncStatus.RPC_S_OK, status, "Complete asynchronous wait.");
            int returnValue = Marshal.ReadInt32(reply);
            flagsOut = Marshal.ReadInt32(this.pulFlagsOut);
            this.VerifyAsyncCallOnRPCTransport();
            return returnValue;
        }

        /// <summary>
        /// Create a subscription for specified notifications on the server. 
        /// </summary>
        /// <param name="notificationType">The notification type which want to subscribe.</param>
        /// <returns>The server response.</returns>
        public RopRegisterNotificationResponse RegisterNotification(NotificationType notificationType)
        {
            RopRegisterNotificationRequest registerNotificationRequest;
            RopRegisterNotificationResponse registerNotificationResponse;

            registerNotificationRequest.RopId = (byte)RopId.RopRegisterNotification;
            byte logonId = this.LogonId;
            registerNotificationRequest.LogonId = logonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            registerNotificationRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x01, which specifies the location in the Server object handle table where the handle,
            // for the output Server object will be stored, as specified in [MS-OXCROPS] section 2.2.6.2.1
            registerNotificationRequest.OutputHandleIndex = 1;

            // Set the type of the notification that the client is interested in receiving.
            registerNotificationRequest.NotificationTypes = (byte)notificationType;

            // This field is reserved. The field value MUST be 0x00.
            registerNotificationRequest.Reserved = 0;

            // TRUE: the scope for notifications is the entire database
            registerNotificationRequest.WantWholeStore = 1;

            // MessageId and FolderId are not present when the value of the WantWholeStore field is nonzero.
            registerNotificationRequest.MessageId = 0;
            registerNotificationRequest.FolderId = 0;

            IList<IDeserializable> responseMessages = this.Process(
                registerNotificationRequest,
                this.LogonHandle,
                out this.responseSOHs);
            registerNotificationResponse = (RopRegisterNotificationResponse)responseMessages[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
            uint registerNotificationResponseHandle = this.responseSOHs[0][registerNotificationResponse.OutputHandleIndex];
            this.VerifyRopRegisterNotificationResponseHandle(registerNotificationResponseHandle);
            this.VerifyROPTransport();
            this.VerifyMAPITransport();
            return registerNotificationResponse;
        }

        /// <summary>
        /// Creates a subscription for specified notifications on the server. 
        /// </summary>
        /// <param name="notificationType">The notification type which want to subscribe.</param>
        /// <param name="wantWholeStore">The value of WantWholeStore.</param>
        /// <param name="flolderId">The value of specified folder ID.</param>
        /// <param name="messageId">The value of specified message ID</param>
        /// <param name="notificationHandle">The value of notificationHandle</param>
        /// <returns>The server response.</returns>
        public RopRegisterNotificationResponse RegisterNotificationWithParameter(NotificationType notificationType, byte wantWholeStore, ulong flolderId, ulong messageId, out uint notificationHandle)
        {
            RopRegisterNotificationRequest registerNotificationRequest;
            RopRegisterNotificationResponse registerNotificationResponse;

            registerNotificationRequest.RopId = (byte)RopId.RopRegisterNotification;
            byte logonId = this.LogonId;
            registerNotificationRequest.LogonId = logonId;

            // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table
            // where the handle for the input Server object is stored.
            registerNotificationRequest.InputHandleIndex = 0;

            // Set OutputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle,
            // for the output Server object will be stored, as specified in [MS-OXCROPS] section 2.2.6.2.1
            registerNotificationRequest.OutputHandleIndex = 1;

            // Set the type of the notification that the client is interested in receiving.
            registerNotificationRequest.NotificationTypes = (byte)notificationType;

            // This field is reserved. The field value MUST be 0x00.
            registerNotificationRequest.Reserved = 0;

            // If the scope for notifications is the entire database, the value of wantWholeStore is true; otherwise, FALSE (0x00).
            registerNotificationRequest.WantWholeStore = wantWholeStore;

            // Set the value of the specified message ID.
            registerNotificationRequest.MessageId = messageId;

            // Set the value of the specified folder ID.
            registerNotificationRequest.FolderId = flolderId;

            IList<IDeserializable> responseMessages = this.Process(
                registerNotificationRequest,
                this.LogonHandle,
                out this.responseSOHs);
            registerNotificationResponse = (RopRegisterNotificationResponse)responseMessages[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                registerNotificationResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
            notificationHandle = this.responseSOHs[0][registerNotificationResponse.OutputHandleIndex];
            this.VerifyRopRegisterNotificationResponseHandle(notificationHandle);
            this.VerifyROPTransport();
            this.VerifyMAPITransport();
            return registerNotificationResponse;
        }

        /// <summary>
        /// Retrieve the RopNotify and RopPending response by sending an empty RPC request.
        /// </summary>
        /// <param name="expectGetNotification">A bool type indicating whether expect to get any notification.</param>
        /// <returns>The list of ROP response, empty if no response returned.</returns>
        public IList<IDeserializable> GetNotification(bool expectGetNotification)
        {
            IList<IDeserializable> notification;
            List<List<uint>> responseSOHs;

            // If expectGetNotification is true, which means the notifications are expected. Do a loop to get notifications.
            // Otherwise, only get the notification once.
            if (expectGetNotification)
            {
                // The retry times to try getting notifications.
                int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
                int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
                do
                {
                    // Sleep some time to get pending notification and notification details,
                    // the sleep time is implementation-specific, which can be configured.
                    Thread.Sleep(sleepTime);

                    notification = this.Process(
                        null,
                        this.LogonHandle,
                        out responseSOHs);
                    retryCount--;
                }
                while (notification.Count == 0 && retryCount > 0);
                Site.Assert.AreNotEqual<int>(
                    0,
                    notification.Count,
                    "Failed to get the RopNotify and RopPending response.");
                foreach (IDeserializable rsp in notification)
                {
                    if (rsp is RopNotifyResponse)
                    {
                        RopNotifyResponse response = (RopNotifyResponse)rsp;
                        this.VerifyRopNotifyResponse(response);
                    }
                    else if (rsp is RopPendingResponse)
                    {
                        this.VerifyRopPendingResponse();
                    }

                    this.VerifyROPTransport();
                    this.VerifyMAPITransport();
                }
            }
            else
            {
                notification = this.Process(
                        null,
                        this.LogonHandle,
                        out responseSOHs);
            }

            return notification;
        }

        /// <summary>
        /// Send ROP request with single operation, single input object handle.
        /// </summary>
        /// <param name="requestRop">The ROP request.</param>
        /// <param name="inputObjHandles">The input object handle.</param>
        /// <param name="responseSOHs">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <returns>The responses returned by the server.</returns>
        public IList<IDeserializable> Process(ISerializable requestRop, uint inputObjHandles, out List<List<uint>> responseSOHs)
        {
            return this.Process(requestRop, new uint[] { inputObjHandles }, out responseSOHs);
        }

        /// <summary>
        /// Send ROP request with single operation, multiple input object handles.
        /// </summary>
        /// <param name="ropRequest">The ROP request.</param>
        /// <param name="inputObjHandles">The multiple input object handles.</param>
        /// <param name="responseSOHs">The Response SOH Table returned by ROP call, provides information like object handle.</param>
        /// <returns>The responses returned by the server.</returns>
        public IList<IDeserializable> Process(ISerializable ropRequest, IEnumerable<uint> inputObjHandles, out List<List<uint>> responseSOHs)
        {
            List<ISerializable> requestRops = null;
            if (ropRequest != null)
            {
                requestRops = new List<ISerializable>
                {
                    ropRequest
                };
            }

            List<uint> requestSOH = new List<uint>();
            if (inputObjHandles != null)
            {
                requestSOH.AddRange(inputObjHandles);

                if (ropRequest != null && Common.IsOutputHandleInRopRequest(ropRequest))
                {
                    // Add an element for server output object handle, set default value to 0xFFFFFFFF
                    requestSOH.Add(0xFFFFFFFF);
                }
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();

            byte[] rawData = new byte[10008];
            responseSOHs = new List<List<uint>>();
            uint length = 0;

            if (Common.IsRequirementEnabled(81001, this.Site))
            {
                // Refer to MS-OXCRPC section 3.1.4.2 and endnote 22, the minimal value of pcbOut is 0x00008007 in Exchange 2007. 
                // So set the value of pcbOut to 0xC350 in Exchange 2007.
                length = 0xC350;
            }

            if (Common.IsRequirementEnabled(81002, this.Site))
            {
                // Refer to MS-OXCRPC section 3.1.4.2 and endnote 22, the minimal value of pcbOut is 0x00000008 in Exchange 2010 and Exchange 2013.
                // Set the value of pcbOut to 0x190 in Exchange 2010 and Exchange 2013.
                length = 0x190;
            }

            uint maxPcbout = (uint)(ropRequest == null ? length : 0x10008);
            uint ret = this.oxcRopsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, maxPcbout);
            if (ret == OxcRpcErrorCode.ECRpcFormat)
            {
                this.Site.Assert.Fail("Error RPC Format");
            }

            this.Site.Assert.AreEqual<uint>(0x0, ret, "If the response is success, the return value is 0x0.");

            return responseRops.ToList();
        }

        /// <summary>
        /// Logon to the server
        /// </summary>
        /// <returns>Server object handle of logon</returns>
        public RopLogonResponse Logon()
        {
            RopLogonRequest logonRequest;

            logonRequest.RopId = 0xFE;
            logonRequest.LogonId = 0x0;
            logonRequest.OutputHandleIndex = 0x0;

            string userDN = Common.GetConfigurationPropertyValue("User1Essdn", this.Site) + "\0";

            logonRequest.StoreState = 0;

            // logon to a private mailbox
            logonRequest.LogonFlags = 0x01;

            // requesting access to the mail box
            logonRequest.OpenFlags = 0x01000000;
            logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);
            logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);

            Collection<uint> inputObjectHandles = new Collection<uint> { 0 };
            IList<IDeserializable> responses = this.Process(logonRequest, inputObjectHandles, out this.responseSOHs);

            RopLogonResponse logonResponse = (RopLogonResponse)responses[0];
            this.Site.Assert.AreEqual<uint>(
                0,
                logonResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success).");
            this.LogonHandle = this.responseSOHs[0][logonResponse.OutputHandleIndex];
            return logonResponse;
        }

        /// <summary>
        /// The method to send NotificationWait request to the server.
        /// </summary>        
        /// <param name="requestBody">The NotificationWait request body.</param>
        /// <returns>Return the NotificationWait response from the server.</returns>
        public NotificationWaitSuccessResponseBody NotificationWait(NotificationWaitRequestBody requestBody)
        {
            return this.oxcRopsClient.MAPINotificationWaitCall(requestBody);
        }

        /// <summary>
        /// Check whether the array of byte is null terminated ASCII string
        /// </summary>
        /// <param name="buffer">The array of byte which to be checked whether  null terminated ASCII string</param>
        /// <returns>A Boolean type indicating whether the passed string is a null-terminated ASCII string.</returns>
        private bool IsNullTerminatedASCIIStr(byte[] buffer)
        {
            int len = buffer.Length;
            bool isNullTerminated = buffer[len - 1] == 0x00;
            bool isASCIIString = true;
            for (int i = 0; i < buffer.Length; i++)
            {
                // ASCII between 0x00 and 0x7F.
                if (buffer[i] <= 0x7F)
                {
                    continue;
                }
                else
                {
                    isASCIIString = false;
                    break;
                }
            }

            bool isNullTerminatedASCIIStr = isNullTerminated && isASCIIString;
            return isNullTerminatedASCIIStr;
        }
    }
}