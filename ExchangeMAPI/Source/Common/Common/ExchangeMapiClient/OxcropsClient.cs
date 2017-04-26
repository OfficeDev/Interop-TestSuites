namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that implements ROP parsing and supports ROP buffer transmission between client and server.
    /// </summary>
    public class OxcropsClient
    {
        #region private Fields

        /// <summary>
        /// Length of RPC_HEADER_EXT
        /// </summary>
        private static readonly int RPCHEADEREXTLEN = Marshal.SizeOf(typeof(RPC_HEADER_EXT));

        /// <summary>
        /// Status of connection.
        /// </summary>
        private bool isConnected;

        /// <summary>
        /// An instance of ITestSite
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// RPC session context handle used by a client when issuing RPC calls against a server. 
        /// </summary>
        private IntPtr cxh;

        /// <summary>
        /// The user name
        /// </summary>
        private string userName;

        /// <summary>
        /// The user DN
        /// </summary>
        private string userDN;

        /// <summary>
        /// User password
        /// </summary>
        private string userPassword;

        /// <summary>
        /// Domain name
        /// </summary>
        private string domainName;

        /// <summary>
        /// Original server name
        /// </summary>
        private string originalServerName;

        /// <summary>
        /// Reserved rop id hash table
        /// </summary>
        private Hashtable reservedRopIdsHT = null;

        /// <summary>
        /// Server name for logon public folder
        /// </summary>
        private string publicFolderServer = null;

        /// <summary>
        /// Server name for logon private mailbox
        /// </summary>
        private string privateMailboxServer = null;

        /// <summary>
        /// Proxy name for logon public folder
        /// </summary>
        private string publicFolderProxyServer = null;

        /// <summary>
        /// Proxy name for logon private mailbox
        /// </summary>
        private string privateMailboxProxyServer = null;

        /// <summary>
        /// Mail store url for private mailbox
        /// </summary>
        private string privateMailStoreUrl = null;

        /// <summary>
        /// Mail store url for public folder
        /// </summary>
        private string publicFolderUrl = null;

        /// <summary>
        /// RpcAdapter instance
        /// </summary>
        private RpcAdapter rpcAdapter;

        /// <summary>
        /// MapiHttpAdapter instance
        /// </summary>
        private MapiHttpAdapter mapiHttpAdapter;

        #endregion

        /// <summary>
        /// Initializes a new instance of the OxcropsClient class.
        /// </summary>
        public OxcropsClient()
        {
        }

        /// <summary>
        /// Initializes a new instance of the OxcropsClient class.
        /// </summary>
        /// <param name="mapiContext">The Mapi Context</param>
        public OxcropsClient(MapiContext mapiContext)
        {
            if (mapiContext == null)
            {
                throw new ArgumentNullException("mapiContext should not be null");
            }

            this.RegisterROPDeserializer();
            this.MapiContext = mapiContext;
            this.site = mapiContext.TestSite;

            switch (mapiContext.TransportSequence.ToLower())
            {
                case "mapi_http":
                    this.mapiHttpAdapter = new MapiHttpAdapter(this.site);
                    break;
                case "ncacn_ip_tcp":
                case "ncacn_http":
                    this.rpcAdapter = new RpcAdapter(this.site);
                    break;
                default:
                    this.site.Assert.Fail("TransportSeq \"{0}\" is not supported by the test suite.");
                    break;
            }
        }

        /// <summary>
        /// Gets or sets mapi context
        /// </summary>
        public MapiContext MapiContext { get; set; }

        /// <summary>
        /// Gets or sets the RPC Context Pointer.
        /// </summary>
        public IntPtr CXH
        {
            get
            {
                return this.cxh;
            }

            set
            {
                this.cxh = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the status of connection is connected.
        /// </summary>
        public bool IsConnected
        {
            get
            {
                return this.isConnected;
            }

            set
            {
                this.isConnected = value;
            }
        }

        /// <summary>
        /// Check RopId whether in Reserved RopIds array
        /// </summary>
        /// <param name="ropId">The RopId will be checked</param>
        /// <returns>If the RopId is a reserved RopId, return true, else return false</returns>
        public bool IsReservedRopId(byte ropId)
        {
            if (this.reservedRopIdsHT == null)
            {
                this.reservedRopIdsHT = new Hashtable();

                // Reserved RopId array
                byte[] reservedRopIds = 
                {
                    0x00, 0x28, 0x3C, 0x3D, 0x62, 0x65, 0x6A, 0x71, 0x7C, 0x7D,
                    0x85, 0x87, 0x8A, 0x8B, 0x8C, 0x8D, 0x8E, 0x94, 0x95, 0x96,
                    0x97, 0x98, 0x99, 0x9A, 0x9B, 0x9C, 0x9D, 0x9E, 0x9F, 0xA0,
                    0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA,
                    0xAB, 0xAC, 0xAD, 0xAE, 0xAF, 0xB0, 0xB1, 0xB2, 0xB3, 0xB4,
                    0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD, 0xBE,
                    0xBF, 0xC0, 0xC1, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8,
                    0xC9, 0xCA, 0xCB, 0xCC, 0xCD, 0xCE, 0xCF, 0xD0, 0xD1, 0xD2,
                    0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xDB, 0xDC,
                    0xDD, 0xDE, 0xDF, 0xE0, 0xE1, 0xE2, 0xE3, 0xE4, 0xE5, 0xE6,
                    0xE7, 0xE8, 0xE9, 0xEA, 0xEB, 0xEC, 0xED, 0xEE, 0xEF, 0xF0, 
                    0xF1, 0xF2, 0xF3, 0xF4, 0xF5, 0xF6, 0xF7, 0xF8, 0xFA, 0xFB,
                    0xFC, 0xFD, 0x52, 0x83, 0x84, 0x88
                };

                for (int i = 0; i < reservedRopIds.Length; i++)
                {
                    this.reservedRopIdsHT.Add(reservedRopIds[i], "ROPID" + i.ToString());
                }
            }

            // Check ropId whether in Reserved RopIds hashtable
            return this.reservedRopIdsHT.ContainsKey(ropId);
        }

        /// <summary>
        /// Connect to the server for running ROP commands.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">the type of connection</param>
        /// <param name="userDN">UserDN used to connect server</param>
        /// <param name="domain">Domain name</param>
        /// <param name="userName">User name used to logon</param>
        /// <param name="password">User Password</param>
        /// <returns>Result of connecting.</returns>
        public bool Connect(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password)
        {
            this.privateMailboxServer = null;
            this.privateMailboxProxyServer = null;
            this.publicFolderServer = null;
            this.publicFolderProxyServer = null;
            this.privateMailStoreUrl = null;
            this.publicFolderUrl = null;

            this.userName = userName;
            this.userDN = userDN;
            this.userPassword = password;
            this.domainName = domain;
            this.originalServerName = server;

            if ((this.MapiContext.AutoRedirect == true) && (Common.GetConfigurationPropertyValue("UseAutodiscover", this.site).ToLower() == "true"))
            {
                string requestURL = Common.GetConfigurationPropertyValue("AutoDiscoverUrlFormat", this.site);
                requestURL = Regex.Replace(requestURL, @"\[ServerName\]", this.originalServerName, RegexOptions.IgnoreCase);
                AutoDiscoverProperties autoDiscoverProperties = AutoDiscover.GetAutoDiscoverProperties(this.site, this.originalServerName, this.userName, this.domainName, requestURL, this.MapiContext.TransportSequence.ToLower());

                this.privateMailboxServer = autoDiscoverProperties.PrivateMailboxServer;
                this.privateMailboxProxyServer = autoDiscoverProperties.PrivateMailboxProxy;
                this.publicFolderServer = autoDiscoverProperties.PublicMailboxServer;
                this.publicFolderProxyServer = autoDiscoverProperties.PublicMailboxProxy;
                this.privateMailStoreUrl = autoDiscoverProperties.PrivateMailStoreUrl;
                this.publicFolderUrl = autoDiscoverProperties.PublicMailStoreUrl;
            }
            else
            {
                if (this.MapiContext.TransportSequence.ToLower() == "mapi_http")
                {
                    this.site.Assert.Fail("When the value of TransportSeq is set to mapi_http, the value of UseAutodiscover must be set to true.");
                }
                else
                {
                    this.publicFolderServer = server;
                    this.privateMailboxServer = server;
                }
            }

            bool ret = false;

            switch (this.MapiContext.TransportSequence.ToLower())
            {
                case "mapi_http":
                    if (connectionType == ConnectionType.PrivateMailboxServer)
                    {
                        ret = this.MapiConnect(this.privateMailStoreUrl, userDN, domain, userName, password);
                    }
                    else
                    {
                        ret = this.MapiConnect(this.publicFolderUrl, userDN, domain, userName, password);
                    }

                    break;
                case "ncacn_ip_tcp":
                case "ncacn_http":
                    if (connectionType == ConnectionType.PrivateMailboxServer)
                    {
                        ret = this.RpcConnect(this.privateMailboxServer, userDN, domain, userName, password, this.privateMailboxProxyServer);
                    }
                    else
                    {
                        ret = this.RpcConnect(this.publicFolderServer, userDN, domain, userName, password, this.publicFolderProxyServer);
                    }

                    break;
                default:
                    this.site.Assert.Fail("TransportSeq \"{0}\" is not supported by the test suite.");
                    break;
            }

            return ret;
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>Result of disconnecting.</returns>
        public bool Disconnect()
        {
            uint ret = 0;

            this.privateMailboxServer = null;
            this.privateMailboxProxyServer = null;

            this.publicFolderServer = null;
            this.publicFolderProxyServer = null;

            this.privateMailStoreUrl = null;
            this.publicFolderUrl = null;

            if (this.IsConnected)
            {
                switch (this.MapiContext.TransportSequence.ToLower())
                {
                    case "mapi_http":
                        ret = this.mapiHttpAdapter.Disconnect();
                        break;
                    case "ncacn_ip_tcp":
                    case "ncacn_http":
                        ret = this.rpcAdapter.Disconnect(ref this.cxh);
                        break;
                    default:
                        this.site.Assert.Fail("TransportSeq \"{0}\" is not supported by the test suite.");
                        break;
                }

                if (ret == 0)
                {
                    this.IsConnected = false;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Send ROP request to the server.
        /// </summary>
        /// <param name="requestROPs">ROP request objects.</param>
        /// <param name="requestSOHTable">ROP request server object handle table.</param>
        /// <param name="responseROPs">ROP response objects.</param>
        /// <param name="responseSOHTable">ROP response server object handle table.</param>
        /// <param name="rgbRopOut">The response payload bytes.</param>
        /// <param name="pcbOut">The maximum size of the rgbOut buffer to place Response in.</param>
        /// <param name="mailBoxUserName">Autodiscover find the mailbox according to this username.</param>
        /// <returns>0 indicates success, other values indicate failure. </returns>
        public uint RopCall(
            List<ISerializable> requestROPs,
            List<uint> requestSOHTable,
            ref List<IDeserializable> responseROPs,
            ref List<List<uint>> responseSOHTable,
            ref byte[] rgbRopOut,
            uint pcbOut,
            string mailBoxUserName = null)
        {
            // Log the rop requests
            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    byte[] ropData = requestROP.Serialize();
                    this.site.Log.Add(LogEntryKind.Comment, "Request: {0}", requestROP.GetType().Name);
                    this.site.Log.Add(LogEntryKind.Comment, Common.FormatBinaryDate(ropData));
                }
            }

            // Construct request buffer
            byte[] rgbIn = this.BuildRequestBuffer(requestROPs, requestSOHTable);

            uint ret = 0;
            switch (this.MapiContext.TransportSequence.ToLower())
            {
                case "mapi_http":
                    ret = this.mapiHttpAdapter.Execute(rgbIn, pcbOut, out rgbRopOut);
                    break;
                case "ncacn_ip_tcp":
                case "ncacn_http":
                    ret = this.rpcAdapter.RpcExt2(ref this.cxh, rgbIn, out rgbRopOut, ref pcbOut);
                    break;
                default:
                    this.site.Assert.Fail("TransportSeq \"{0}\" is not supported by the test suite.");
                    break;
            }

            RPC_HEADER_EXT[] rpcHeaderExts;
            byte[][] rops;
            uint[][] serverHandleObjectsTables;

            if (ret == OxcRpcErrorCode.ECNone)
            {
                this.ParseResponseBuffer(rgbRopOut, out rpcHeaderExts, out rops, out serverHandleObjectsTables);

                // Deserialize rops
                if (rops != null)
                {
                    foreach (byte[] rop in rops)
                    {
                        List<IDeserializable> ropResponses = new List<IDeserializable>();
                        RopDeserializer.Deserialize(rop, ref ropResponses);
                        foreach (IDeserializable ropResponse in ropResponses)
                        {
                            responseROPs.Add(ropResponse);
                            Type type = ropResponse.GetType();
                            this.site.Log.Add(LogEntryKind.Comment, "Response: {0}", type.Name);
                        }

                        this.site.Log.Add(LogEntryKind.Comment, Common.FormatBinaryDate(rop));
                    }
                }

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

                // The return value 0x478 means that the client needs to reconnect server with server name in response 
                if (this.MapiContext.AutoRedirect && rops.Length > 0 && rops[0][0] == 0xfe && ((RopLogonResponse)responseROPs[0]).ReturnValue == 0x478)
                {
                    // Reconnect server with returned server name
                    string serverName = Encoding.ASCII.GetString(((RopLogonResponse)responseROPs[0]).ServerName);
                    serverName = serverName.Substring(serverName.LastIndexOf("=") + 1);

                    responseROPs.Clear();
                    responseSOHTable.Clear();

                    bool disconnectReturnValue = this.Disconnect();
                    this.site.Assert.IsTrue(disconnectReturnValue, "Disconnect should be successful here.");

                    string rpcProxyOptions = null;
                    if (string.Compare(this.MapiContext.TransportSequence, "ncacn_http", true) == 0)
                    {
                        rpcProxyOptions = "RpcProxy=" + this.originalServerName + "." + this.domainName;

                        bool connectionReturnValue = this.RpcConnect(serverName, this.userDN, this.domainName, this.userName, this.userPassword, rpcProxyOptions);
                        this.site.Assert.IsTrue(connectionReturnValue, "RpcConnect_Internal should be successful here.");
                    }
                    else if (string.Compare(this.MapiContext.TransportSequence, "mapi_http", true) == 0)
                    {
                        if (mailBoxUserName == null)
                        {
                            mailBoxUserName = Common.GetConfigurationPropertyValue("AdminUserName", this.site);
                            if (mailBoxUserName == null || mailBoxUserName == "")
                            {
                                this.site.Assert.Fail(@"There must be ""AdminUserName"" configure item in the ptfconfig file.");
                            }
                        }

                        string requestURL = Common.GetConfigurationPropertyValue("AutoDiscoverUrlFormat", this.site);                        
                        requestURL = Regex.Replace(requestURL, @"\[ServerName\]", this.originalServerName, RegexOptions.IgnoreCase);
                        AutoDiscoverProperties autoDiscoverProperties = AutoDiscover.GetAutoDiscoverProperties(this.site, this.originalServerName, mailBoxUserName, this.domainName, requestURL, this.MapiContext.TransportSequence.ToLower());

                        this.privateMailboxServer = autoDiscoverProperties.PrivateMailboxServer;
                        this.privateMailboxProxyServer = autoDiscoverProperties.PrivateMailboxProxy;
                        this.publicFolderServer = autoDiscoverProperties.PublicMailboxServer;
                        this.publicFolderProxyServer = autoDiscoverProperties.PublicMailboxProxy;
                        this.privateMailStoreUrl = autoDiscoverProperties.PrivateMailStoreUrl;
                        this.publicFolderUrl = autoDiscoverProperties.PublicMailStoreUrl;

                        bool connectionReturnValue = this.MapiConnect(this.privateMailStoreUrl, this.userDN, this.domainName, this.userName, this.userPassword);
                        this.site.Assert.IsTrue(connectionReturnValue, "RpcConnect_Internal should be successful here.");
                    }                    

                    ret = this.RopCall(
                        requestROPs,
                        requestSOHTable,
                        ref responseROPs,
                        ref responseSOHTable,
                        ref rgbRopOut,
                        0x10008);
                }
            }

            return ret;
        }

        /// <summary>
        /// The method to send NotificationWait request to the server.
        /// </summary>        
        /// <param name="requestBody">The NotificationWait request body.</param>
        /// <returns>Return the NotificationWait response body.</returns>
        public NotificationWaitSuccessResponseBody MAPINotificationWaitCall(IRequestBody requestBody)
        {
            return this.mapiHttpAdapter.NotificationWaitCall(requestBody);
        }

        #region Register ROPs' deserializer
        /// <summary>
        /// Register ROPs' deserializer
        /// </summary>
        public void RegisterROPDeserializer()
        {
            RopDeserializer.Init();

            // Logon ROPs response register
            this.RegisterLogonROPDeserializer();

            // Fast Transfer ROPs response register
            this.RegisterFastTransferROPDeserializer();

            // Folder ROPs response register
            this.RegisterFolderROPDeserializer();

            // Incremental Change Synchronization ROPs response register
            this.RegisterIncrementalChangeSynchronizationROPDeserializer();

            // Message ROPs response register
            this.RegisterMessageROPDeserializer();

            // Notification ROPs response register
            this.RegisterNotificationROPDeserializer();

            // Other ROPs response register
            this.RegisterOtherROPDeserializer();

            // Permission ROPs response register
            this.RegisterPermissionROPDeserializer();

            // Property ROPs
            this.RegisterPropertyROPDeserializer();

            // Rule ROPs response register
            this.RegisterRuleROPDeserializer();

            // Stream ROPs response register
            this.RegisterStreamROPDeserializer();

            // Table ROPs response register
            this.RegisterTableROPDeserializer();

            // Transport ROPs response register
            this.RegisterTransportROPDeserializer();
        }

        /// <summary>
        /// Register Logon ROPs' deserializer
        /// </summary>
        private void RegisterLogonROPDeserializer()
        {
            #region Logon ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetOwningServers), new RopGetOwningServersResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPerUserGuid), new RopGetPerUserGuidResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPerUserLongTermIds), new RopGetPerUserLongTermIdsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetReceiveFolder), new RopGetReceiveFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetReceiveFolderTable), new RopGetReceiveFolderTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetStoreState), new RopGetStoreStateResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopIdFromLongTermId), new RopIdFromLongTermIdResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopLogon), new RopLogonResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopLongTermIdFromId), new RopLongTermIdFromIdResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenAttachment), new RopOpenAttachmentResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopPublicFolderIsGhosted), new RopPublicFolderIsGhostedResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopReadPerUserInformation), new RopReadPerUserInformationResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetReceiveFolder), new RopSetReceiveFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopWritePerUserInformation), new RopWritePerUserInformationResponse());
            #endregion
        }

        /// <summary>
        /// Register Fast Transfer ROPs' deserializer
        /// </summary>
        private void RegisterFastTransferROPDeserializer()
        {
            #region Fast Transfer ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferDestinationConfigure), new RopFastTransferDestinationConfigureResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferDestinationPutBuffer), new RopFastTransferDestinationPutBufferResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferSourceCopyFolder), new RopFastTransferSourceCopyFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferSourceCopyMessages), new RopFastTransferSourceCopyMessagesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferSourceCopyProperties), new RopFastTransferSourceCopyPropertiesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferSourceCopyTo), new RopFastTransferSourceCopyToResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFastTransferSourceGetBuffer), new RopFastTransferSourceGetBufferResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopTellVersion), new RopTellVersionResponse());
            #endregion
        }

        /// <summary>
        /// Register Folder ROPs' deserializer
        /// </summary>
        private void RegisterFolderROPDeserializer()
        {
            #region Folder ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCopyFolder), new RopCopyFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCreateFolder), new RopCreateFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopDeleteFolder), new RopDeleteFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopDeleteMessages), new RopDeleteMessagesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopEmptyFolder), new RopEmptyFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetContentsTable), new RopGetContentsTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetHierarchyTable), new RopGetHierarchyTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetSearchCriteria), new RopGetSearchCriteriaResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopHardDeleteMessages), new RopHardDeleteMessagesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopHardDeleteMessagesAndSubfolders), new RopHardDeleteMessagesAndSubfoldersResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopMoveCopyMessages), new RopMoveCopyMessagesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopMoveFolder), new RopMoveFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenFolder), new RopOpenFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetSearchCriteria), new RopSetSearchCriteriaResponse());
            #endregion
        }

        /// <summary>
        /// Register Incremental Change Synchronization ROPs' deserializer
        /// </summary>
        private void RegisterIncrementalChangeSynchronizationROPDeserializer()
        {
            #region Incremental Change Synchronization ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetLocalReplicaIds), new RopGetLocalReplicaIdsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetLocalReplicaMidsetDeleted), new RopSetLocalReplicaMidsetDeletedResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationConfigure), new RopSynchronizationConfigureResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationGetTransferState), new RopSynchronizationGetTransferStateResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationImportDeletes), new RopSynchronizationImportDeletesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationImportHierarchyChange), new RopSynchronizationImportHierarchyChangeResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationImportMessageChange), new RopSynchronizationImportMessageChangeResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationImportMessageMove), new RopSynchronizationImportMessageMoveResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationImportReadStateChanges), new RopSynchronizationImportReadStateChangesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationOpenCollector), new RopSynchronizationOpenCollectorResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationUploadStateStreamBegin), new RopSynchronizationUploadStateStreamBeginResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationUploadStateStreamContinue), new RopSynchronizationUploadStateStreamContinueResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSynchronizationUploadStateStreamEnd), new RopSynchronizationUploadStateStreamEndResponse());
            #endregion
        }

        /// <summary>
        /// Register Message ROPs' deserializer
        /// </summary>
        private void RegisterMessageROPDeserializer()
        {
            #region Message ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCreateAttachment), new RopCreateAttachmentResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCreateMessage), new RopCreateMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopDeleteAttachment), new RopDeleteAttachmentResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetAttachmentTable), new RopGetAttachmentTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetMessageStatus), new RopGetMessageStatusResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopModifyRecipients), new RopModifyRecipientsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenEmbeddedMessage), new RopOpenEmbeddedMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenMessage), new RopOpenMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopReadRecipients), new RopReadRecipientsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopReloadCachedInformation), new RopReloadCachedInformationResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopRemoveAllRecipients), new RopRemoveAllRecipientsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSaveChangesAttachment), new RopSaveChangesAttachmentResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSaveChangesMessage), new RopSaveChangesMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetMessageReadFlag), new RopSetMessageReadFlagResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetMessageStatus), new RopSetMessageStatusResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetReadFlags), new RopSetReadFlagsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetValidAttachments), new RopGetValidAttachmentsResponse());
            #endregion
        }

        /// <summary>
        /// Register Notification ROPs' deserializer
        /// </summary>
        private void RegisterNotificationROPDeserializer()
        {
            #region Notification ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopNotify), new RopNotifyResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopPending), new RopPendingResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopRegisterNotification), new RopRegisterNotificationResponse());
            #endregion
        }

        /// <summary>
        /// Register other ROPs' deserializer
        /// </summary>
        private void RegisterOtherROPDeserializer()
        {
            #region Other ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopBackoff), new RopBackoffResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopBufferTooSmall), new RopBufferTooSmallResponse());
            #endregion
        }

        /// <summary>
        /// Register Permission ROPs' deserializer
        /// </summary>
        private void RegisterPermissionROPDeserializer()
        {
            #region Permission ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPermissionsTable), new RopGetPermissionsTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopModifyPermissions), new RopModifyPermissionsResponse());
            #endregion
        }

        /// <summary>
        /// Register Property ROPs' deserializer
        /// </summary>
        private void RegisterPropertyROPDeserializer()
        {
            #region Property ROPs
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenAttachment), new RopOpenAttachmentResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCopyProperties), new RopCopyPropertiesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCopyTo), new RopCopyToResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopDeleteProperties), new RopDeletePropertiesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopDeletePropertiesNoReplicate), new RopDeletePropertiesNoReplicateResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetNamesFromPropertyIds), new RopGetNamesFromPropertyIdsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPropertiesAll), new RopGetPropertiesAllResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPropertiesList), new RopGetPropertiesListResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPropertiesSpecific), new RopGetPropertiesSpecificResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetPropertyIdsFromNames), new RopGetPropertyIdsFromNamesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopProgress), new RopProgressResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopQueryNamedProperties), new RopQueryNamedPropertiesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetProperties), new RopSetPropertiesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetPropertiesNoReplicate), new RopSetPropertiesNoReplicateResponse());
            #endregion
        }

        /// <summary>
        /// Register Rule ROPs' deserializer
        /// </summary>
        private void RegisterRuleROPDeserializer()
        {
            #region Rule ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetRulesTable), new RopGetRulesTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopModifyRules), new RopModifyRulesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopUpdateDeferredActionMessages), new RopUpdateDeferredActionMessagesResponse());
            #endregion
        }

        /// <summary>
        /// Register Stream ROPs' deserializer
        /// </summary>
        private void RegisterStreamROPDeserializer()
        {
            #region Stream ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCommitStream), new RopCommitStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCopyToStream), new RopCopyToStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetStreamSize), new RopGetStreamSizeResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopLockRegionStream), new RopLockRegionStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOpenStream), new RopOpenStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopReadStream), new RopReadStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSeekStream), new RopSeekStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetStreamSize), new RopSetStreamSizeResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopUnlockRegionStream), new RopUnlockRegionStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopWriteStream), new RopWriteStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCloneStream), new RopCloneStreamResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopWriteAndCommitStream), new RopWriteAndCommitStreamResponse());
            #endregion
        }

        /// <summary>
        /// Register Table ROPs' deserializer
        /// </summary>
        private void RegisterTableROPDeserializer()
        {
            #region Table ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopAbort), new RopAbortResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCollapseRow), new RopCollapseRowResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopCreateBookmark), new RopCreateBookmarkResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopExpandRow), new RopExpandRowResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFindRow), new RopFindRowResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopFreeBookmark), new RopFreeBookmarkResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetCollapseState), new RopGetCollapseStateResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetStatus), new RopGetStatusResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopQueryColumnsAll), new RopQueryColumnsAllResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopQueryPosition), new RopQueryPositionResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopQueryRows), new RopQueryRowsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopResetTable), new RopResetTableResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopRestrict), new RopRestrictResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSeekRow), new RopSeekRowResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSeekRowBookmark), new RopSeekRowBookmarkResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSeekRowFractional), new RopSeekRowFractionalResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetCollapseState), new RopSetCollapseStateResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetColumns), new RopSetColumnsResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSortTable), new RopSortTableResponse());
            #endregion
        }

        /// <summary>
        /// Register Transport ROPs' deserializer
        /// </summary>
        private void RegisterTransportROPDeserializer()
        {
            #region Transport ROPs response register
            RopDeserializer.Register(Convert.ToInt32(RopId.RopAbortSubmit), new RopAbortSubmitResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetAddressTypes), new RopGetAddressTypesResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopGetTransportFolder), new RopGetTransportFolderResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopOptionsData), new RopOptionsDataResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSetSpooler), new RopSetSpoolerResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSpoolerLockMessage), new RopSpoolerLockMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopSubmitMessage), new RopSubmitMessageResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopTransportNewMail), new RopTransportNewMailResponse());
            RopDeserializer.Register(Convert.ToInt32(RopId.RopTransportSend), new RopTransportSendResponse());
            #endregion
        }
        #endregion

        /// <summary>
        /// The method parses response buffer.
        /// </summary>
        /// <param name="rgbOut">The ROP response payload.</param>
        /// <param name="rpcHeaderExts">RPC header ext.</param>
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
                // Parse rpc header ext
                RPC_HEADER_EXT rpcHeaderExt;
                ptr = Marshal.AllocHGlobal(RPCHEADEREXTLEN);
                try
                {
                    Marshal.Copy(rgbOut, index, ptr, RPCHEADEREXTLEN);
                    rpcHeaderExt = (RPC_HEADER_EXT)Marshal.PtrToStructure(ptr, typeof(RPC_HEADER_EXT));
                    rpcHeaderExtList.Add(rpcHeaderExt);
                    index += RPCHEADEREXTLEN;
                    end = (rpcHeaderExt.Flags & (ushort)RpcHeaderExtFlags.Last) == (ushort)RpcHeaderExtFlags.Last;
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }

                #region  start parse payload

                // Parse ropSize
                ushort ropSize = BitConverter.ToUInt16(rgbOut, index);
                index += sizeof(ushort);

                if ((ropSize - sizeof(ushort)) > 0)
                {
                    // Parse rop
                    byte[] rop = new byte[ropSize - sizeof(ushort)];
                    Array.Copy(rgbOut, index, rop, 0, ropSize - sizeof(ushort));
                    ropList.Add(rop);
                    index += ropSize - sizeof(ushort);
                }

                // Parse server handle objects table
                this.site.Assert.IsTrue((rpcHeaderExt.Size - ropSize) % sizeof(uint) == 0, "Server object handle should be uint32 array");

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
        /// Create ROPs request buffer.
        /// </summary>
        /// <param name="requestROPs">ROPs in request.</param>
        /// <param name="requestSOHTable">Server object handles table.</param>
        /// <returns>The ROPs request buffer.</returns>
        private byte[] BuildRequestBuffer(List<ISerializable> requestROPs, List<uint> requestSOHTable)
        {
            // Definition for PayloadLen which indicates the length of the field that represents the length of payload.
            int payloadLen = 0x2;
            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    payloadLen += requestROP.Size();
                }
            }

            ushort ropSize = (ushort)payloadLen;

            if (requestSOHTable != null)
            {
                payloadLen += requestSOHTable.Count * sizeof(uint);
            }

            byte[] requestBuffer = new byte[RPCHEADEREXTLEN + payloadLen];
            int index = 0;

            // Construct RPC header ext buffer
            RPC_HEADER_EXT rpcHeaderExt = new RPC_HEADER_EXT
            {
                // There is only one version of the header at this time so this value MUST be set to 0x00.
                Version = 0x00,

                // Last (0x04) indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT. 
                Flags = (ushort)RpcHeaderExtFlags.Last,
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

            // Compress and obfuscate request buffer as configured.
            requestBuffer = Common.CompressAndObfuscateRequest(requestBuffer, this.site);

            return requestBuffer;
        }

        /// <summary>
        /// Internal use connect to the server for RPC calling.
        /// </summary>
        /// <param name="mailStoreUrl">The mail store url used to connect with the server.</param>
        /// <param name="userDN">User DN used to connect server</param>
        /// <param name="domain">Domain name</param>
        /// <param name="userName">User name used to logon.</param>
        /// <param name="password">User Password.</param>
        /// <returns>If client connects server successfully, it returns true, otherwise return false.</returns>
        private bool MapiConnect(string mailStoreUrl, string userDN, string domain, string userName, string password)
        {
            uint returnValue = this.mapiHttpAdapter.Connect(mailStoreUrl, domain, userName, userDN, password);
            if (0 == returnValue)
            {
                this.isConnected = true;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Internal use connect to the server for RPC calling.
        /// This method is defined as a direct way to connect to server with specific parameters.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="userDN">UserDN used to connect server</param>
        /// <param name="domain">Domain name</param>
        /// <param name="userName">User name used to logon</param>
        /// <param name="password">User Password</param>
        /// <param name="rpcProxyOptions">Rpc proxy parameter</param>
        /// <returns>Result of connecting.</returns>
        private bool RpcConnect(string server, string userDN, string domain, string userName, string password, string rpcProxyOptions)
        {
            string userSpn = Regex.Replace(this.MapiContext.SpnFormat, @"\[ServerName\]", this.originalServerName, RegexOptions.IgnoreCase);
            this.cxh = this.rpcAdapter.Connect(server, domain, userName, userDN, password, userSpn, this.MapiContext, rpcProxyOptions);

            if (this.cxh != IntPtr.Zero)
            {
                this.IsConnected = true;
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}