namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Implementation of the MS-OXCSTOR Adapter
    /// </summary>
    public partial class MS_OXCSTORAdapter : ManagedAdapterBase, IMS_OXCSTORAdapter
    {
        #region Variables
        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// Status of connection.
        /// </summary>
        private bool isConnected;

        #endregion Variables

        #region MS_OXCSTORAdapter methods

        /// <summary>
        /// Connect to server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <returns>True indicates connecting successfully, otherwise false</returns>
        public bool ConnectEx(ConnectionType connectionType)
        {
            string domainName = Common.GetConfigurationPropertyValue(ConstValues.Domain, Site);
            string userName = Common.GetConfigurationPropertyValue(ConstValues.UserName, Site);
            string password = Common.GetConfigurationPropertyValue(ConstValues.Password, Site);
            string server = Common.GetConfigurationPropertyValue(ConstValues.Server1, Site);
            string userDN = Common.GetConfigurationPropertyValue(ConstValues.UserEssdn, Site);

            return this.ConnectEx(server, connectionType, userDN, domainName, userName, password);
        }

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
        public bool ConnectEx(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password)
        {
            bool ret = this.oxcropsClient.Connect(server, connectionType, userDN, domain, userName, password);
            this.isConnected = ret;
            return ret;
        }

        /// <summary>
        /// Disconnect the connection with server.
        /// </summary>
        /// <returns>True indicates disconnecting successfully, otherwise false</returns>
        public bool DisconnectEx()
        {
            // Since the operation of receiving ROPs has finished here, this method can be invoked 
            // to verify that "The ROP response buffers are received from the server by using the underlying RPC transport"
            bool ret = this.oxcropsClient.Disconnect();
            if (ret)
            {
                this.isConnected = false;
            }

            return ret;
        }

        /// <summary>
        /// Send ROP request with single operation with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="commandType">ROP commands type</param>
        /// <param name="outputBuffer">ROP response buffer</param>
        /// <param name="mailBoxUser">Mailbox which to logon to</param>
        public void DoRopCall(ISerializable ropRequest, uint inputObjHandle, ROPCommandType commandType, out RopOutputBuffer outputBuffer, string mailBoxUser = null)
        {
            outputBuffer = new RopOutputBuffer();
            List<ISerializable> inputBuffer = new List<ISerializable>
            {
                ropRequest
            };
            List<uint> requestSOH = new List<uint>();
            requestSOH.Add(inputObjHandle);

            if (ropRequest != null && Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responses = new List<IDeserializable>();
            List<List<uint>> responseSOHTable = new List<List<uint>>();
            byte[] rawData = null;
            uint ret = this.oxcropsClient.RopCall(inputBuffer, requestSOH, ref responses, ref responseSOHTable, ref rawData, 0x10008, mailBoxUser);
            if (ret != 0)
            {
                Site.Assert.Fail("Calling RopCall should return 0 for success, but it returns value: {0}", ret);
            }

            this.VerifyRPC();
            this.VerifyTransport();
            outputBuffer.RopsList = responses;
            outputBuffer.ServerObjectHandleTable = responseSOHTable[0];

            IDeserializable response = null;
            if (commandType != ROPCommandType.Others)
            {
                response = responses[0x0];
            }

            switch (commandType)
            {
                case ROPCommandType.RopLogonPrivateMailbox:
                    this.VerifyRopLogonForPrivateMailbox((RopLogonRequest)ropRequest, (RopLogonResponse)response);
                    break;
                case ROPCommandType.RopLogonPublicFolder:
                    this.VerifyRopLogonForPublicFolder((RopLogonRequest)ropRequest, (RopLogonResponse)response);
                    break;
                case ROPCommandType.RopGetOwningServers:
                    this.VerifyRopGetOwningServers((RopGetOwningServersResponse)response);
                    break;
                case ROPCommandType.RopGetPerUserLongTermIds:
                    this.VerifyRopGetPerUserLongTermIds((RopGetPerUserLongTermIdsResponse)response);
                    break;
                case ROPCommandType.RopGetPerUserGuid:
                    this.VerifyRopGetPerUserGuid((RopGetPerUserGuidResponse)response);
                    break;
                case ROPCommandType.RopSetReceiveFolder:
                    this.VerifyRopSetReceiveFolder((RopSetReceiveFolderResponse)response);
                    break;
                case ROPCommandType.RopGetReceiveFolder:
                    this.VerifyRopGetReceiveFolder((RopGetReceiveFolderResponse)response);
                    break;
                case ROPCommandType.RopGetReceiveFolderTable:
                        this.VerifyRopGetReceiveFolderTable((RopGetReceiveFolderTableResponse)response);
                    break;
                case ROPCommandType.RopPublicFolderIsGhosted:
                    this.VerifyRopPublicFolderIsGhosted((RopPublicFolderIsGhostedResponse)response);
                    break;
                case ROPCommandType.RopLongTermIdFromId:
                    this.VerifyRopLongTermIdFromId((RopLongTermIdFromIdRequest)ropRequest, (RopLongTermIdFromIdResponse)response);
                    break;
                case ROPCommandType.RopIdFromLongTermId:
                    this.VerifyRopIdFromLongTermId((RopIdFromLongTermIdRequest)ropRequest, (RopIdFromLongTermIdResponse)response);
                    break;
                case ROPCommandType.RopReadPerUserInformation:
                    this.VerifyRopReadPerUserInformation((RopReadPerUserInformationResponse)response);
                    break;
                case ROPCommandType.RopWritePerUserInformation:
                    this.VerifyRopWritePerUserInformation((RopWritePerUserInformationResponse)response);
                    break;
                case ROPCommandType.Others:
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Set auto redirect value in RPC context
        /// </summary>
        /// <param name="option">True indicates enable auto redirect, false indicates disable auto redirect</param>
        public void SetAutoRedirect(bool option)
        {
            this.oxcropsClient.MapiContext.AutoRedirect = option;
        }

        #endregion MS_OXCSTORAdapter methods

        #region Override functions
        /// <summary>
        /// Initialize class
        /// </summary>
        /// <param name="testSite">The instance of the ITestSite</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCSTOR";
            Common.MergeConfiguration(testSite);
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        /// <summary>
        /// Reset the adapter.
        /// </summary>
        public override void Reset()
        {
            if (this.isConnected)
            {
                this.DisconnectEx();
            }

            base.Reset();
        }

        #endregion Override functions
    }
}