//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Implementation of the MS-OXCFOLD adapter.
    /// </summary>
    public partial class MS_OXCFOLDAdapter : ManagedAdapterBase, IMS_OXCFOLDAdapter
    {
        #region Variables

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// Original bytes array.
        /// </summary>
        private byte[] rawData;
        #endregion Variables

        /// <summary>
        /// Overrides IAdapter's Initialize method, to set testSite.DefaultProtocolDocShortName.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCFOLD";
            Common.MergeConfiguration(testSite);
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        #region MS_OXCFOLDAdapter methods

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <returns>If the behavior of connecting server is successful, the server will return true; otherwise, return false.</returns>
        public bool DoConnect(ConnectionType connectionType)
        {
            return this.oxcropsClient.Connect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    connectionType,
                    Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site),
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                    Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site));
        }

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">the type of connection</param>
        /// <param name="userDN">UserDN used to connect server.</param>
        /// <param name="domain">The domain the server is deployed.</param>
        /// <param name="userName">The domain account name.</param>
        /// <param name="password">Password value.</param>
        /// <returns>If the behavior of connecting server is successful, the server will return true; otherwise, return false.</returns>
        public bool DoConnect(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password)
        {
            return this.oxcropsClient.Connect(
                    server,
                    connectionType,
                    userDN,
                    domain,
                    userName,
                    password);
        }

        /// <summary>
        /// Client calls it to disconnect the connection with Server.
        /// </summary>
        /// <returns>The return value indicates the disconnection state.</returns>
        public bool DoDisconnect()
        {
            return this.oxcropsClient.Disconnect();
        }

        /// <summary>
        /// Sends ROP request with single operation and single input object handle with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="ropResponse">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        public uint DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object ropResponse, ref List<List<uint>> responseSOHTable)
        {
            return this.ExcuteRopCall(ropRequest, insideObjHandle, ref ropResponse, ref responseSOHTable, ref this.rawData);
        }

        /// <summary>
        /// Sends ROP request with single operation and multiple input object handles with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="insideObjHandle">The list of server object handles in request.</param>
        /// <param name="ropResponse">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        public uint DoRopCall(ISerializable ropRequest, List<uint> insideObjHandle, ref object ropResponse, ref List<List<uint>> responseSOHTable)
        {
            return this.ExcuteRopCall(ropRequest, insideObjHandle, ref ropResponse, ref responseSOHTable, ref this.rawData);
        }

        /// <summary>
        /// Creates either public folders or private mailbox folders. 
        /// </summary>
        /// <param name="ropCreateFolderRequest">RopCreateFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopCreateFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopCreateFolderResponse.</param>
        /// <returns>RopCreateFolderResponse object.</returns>
        public RopCreateFolderResponse CreateFolder(RopCreateFolderRequest ropCreateFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropCreateFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopCreateFolderResponse ropCreateFolderResponse = (RopCreateFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropCreateFolderResponse.ReturnValue)
            {
                this.VerifyRopCreateFolder(ropCreateFolderResponse);
            }
            #endregion

            return ropCreateFolderResponse;
        }

        /// <summary>
        /// Opens an existing folder.
        /// </summary>
        /// <param name="ropOpenFolderRequest">RopOpenFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopOpenFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopOpenFolderResponse.</param>
        /// <returns>RopOpenFolderResponse object.</returns>
        public RopOpenFolderResponse OpenFolder(RopOpenFolderRequest ropOpenFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            ropOpenFolderRequest.RopId = (byte)RopId.RopOpenFolder;
            ropOpenFolderRequest.LogonId = Constants.CommonLogonId;
            ropOpenFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            ropOpenFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            this.ExcuteRopCall((ISerializable)ropOpenFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopOpenFolderResponse ropOpenFolderResponse = (RopOpenFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropOpenFolderResponse.ReturnValue)
            {
                this.VerifyRopOpenFolder(ropOpenFolderResponse);
            }
            #endregion

            return ropOpenFolderResponse;
        }

        /// <summary>
        /// Removes a subfolder. 
        /// </summary>
        /// <param name="ropDeleteFolderRequest">RopDeleteFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopDeleteFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopDeleteFolderResponse.</param>
        /// <returns>RopDeleteFolderResponse object.</returns>
        public RopDeleteFolderResponse DeleteFolder(RopDeleteFolderRequest ropDeleteFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropDeleteFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopDeleteFolderResponse ropDeleteFolderResponse = (RopDeleteFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropDeleteFolderResponse.ReturnValue)
            {
                this.VerifyRopDeleteFolder(ropDeleteFolderResponse);
            }
            #endregion

            return ropDeleteFolderResponse;
        }

        /// <summary>
        /// Establishes search criteria for a search folder. 
        /// </summary>
        /// <param name="ropSetSearchCriteriaRequest">RopSetSearchCriteriaRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopSetSearchCriteriaRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopSetSearchCriteriaResponse.</param>
        /// <returns>RopSetSearchCriteriaResponse object.</returns>
        public RopSetSearchCriteriaResponse SetSearchCriteria(RopSetSearchCriteriaRequest ropSetSearchCriteriaRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();

            int count = 0;
            bool setSearchCriteriaComplete = false;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            RopSetSearchCriteriaResponse ropSetSearchCriteriaResponse;
            do
            {
                this.ExcuteRopCall((ISerializable)ropSetSearchCriteriaRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
                ropSetSearchCriteriaResponse = (RopSetSearchCriteriaResponse)temp;

                // Error code 0x499 indicates the server is not ready to initialize a search now, the client should wait for a while and try again later.
                if (0x499 == ropSetSearchCriteriaResponse.ReturnValue)
                {
                    Thread.Sleep(waitTime);
                }
                else
                {
                    setSearchCriteriaComplete = true;
                }

                if (count > retryCount)
                {
                    Site.Assert.Fail("The server failed to initialize a search!");
                }

                count++;
            }
            while (!setSearchCriteriaComplete);

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropSetSearchCriteriaResponse.ReturnValue)
            {
                this.VerifyRopSetSearchCriteria(ropSetSearchCriteriaResponse);
            }
            #endregion Capture Code

            return ropSetSearchCriteriaResponse;
        }

        /// <summary>
        /// Obtains the search criteria and the status of a search for a search folder. 
        /// </summary>
        /// <param name="ropGetSearchCriteriaRequest">RopGetSearchCriteriaRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopGetSearchCriteriaRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetSearchCriteriaResponse.</param>
        /// <returns>RopGetSearchCriteriaResponse object.</returns>
        public RopGetSearchCriteriaResponse GetSearchCriteria(RopGetSearchCriteriaRequest ropGetSearchCriteriaRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropGetSearchCriteriaRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopGetSearchCriteriaResponse ropGetSearchCriteriaResponse = (RopGetSearchCriteriaResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropGetSearchCriteriaResponse.ReturnValue)
            {
                this.VerifyRopGetSearchCriteria(ropGetSearchCriteriaResponse);
            }
            #endregion

            return ropGetSearchCriteriaResponse;
        }

        /// <summary>
        /// Moves or copies messages from a source folder to a destination folder. 
        /// </summary>
        /// <param name="ropMoveCopyMessagesRequest">RopMoveCopyMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopMoveCopyMessagesRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopMoveCopyMessagesResponse.</param>
        /// <returns>RopMoveCopyMessagesResponse object.</returns>
        public RopMoveCopyMessagesResponse MoveCopyMessages(RopMoveCopyMessagesRequest ropMoveCopyMessagesRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropMoveCopyMessagesRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopMoveCopyMessagesResponse ropMoveCopyMessagesResponse = (RopMoveCopyMessagesResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropMoveCopyMessagesResponse.ReturnValue)
            {
                this.VerifyRopMoveCopyMessages(ropMoveCopyMessagesResponse);
            }
            #endregion

            return ropMoveCopyMessagesResponse;
        }

        /// <summary>
        /// Moves a folder from one parent to another.
        /// </summary>
        /// <param name="ropMoveFolderRequest">RopMoveFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopMoveFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopMoveFolderResponse.</param>
        /// <returns>RopMoveFolderResponse object.</returns>
        public RopMoveFolderResponse MoveFolder(RopMoveFolderRequest ropMoveFolderRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropMoveFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopMoveFolderResponse ropMoveFolderResponse = (RopMoveFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropMoveFolderResponse.ReturnValue)
            {
                this.VerifyRopMoveFolder(ropMoveFolderResponse);
            }

            #endregion

            return ropMoveFolderResponse;
        }

        /// <summary>
        /// Creates a new folder on the destination parent folder, copying the properties and content of the source folder to the new folder.
        /// </summary>
        /// <param name="ropCopyFolderRequest">RopCopyFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopCopyFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopCopyFolderResponse.</param>
        /// <returns>RopCopyFolderResponse object.</returns>
        public RopCopyFolderResponse CopyFolder(RopCopyFolderRequest ropCopyFolderRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropCopyFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopCopyFolderResponse ropCopyFolderResponse = (RopCopyFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropCopyFolderResponse.ReturnValue)
            {
                this.VerifyRopCopyFolder(ropCopyFolderResponse);
            }
            #endregion

            return ropCopyFolderResponse;
        }

        /// <summary>
        /// Soft deletes all messages and subfolders from a folder without deleting the folder itself. 
        /// </summary>
        /// <param name="ropEmptyFolderRequest">RopEmptyFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in EmptyFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopEmptyFolderResponse.</param>
        /// <returns>RopEmptyFolderResponse object.</returns>
        public RopEmptyFolderResponse EmptyFolder(RopEmptyFolderRequest ropEmptyFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropEmptyFolderRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopEmptyFolderResponse ropEmptyFolderResponse = (RopEmptyFolderResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropEmptyFolderResponse.ReturnValue)
            {
                this.VerifyRopEmptyFolder(ropEmptyFolderResponse);
            }
            #endregion

            return ropEmptyFolderResponse;
        }

        /// <summary>
        /// Hard deletes all messages and subfolders from a folder without deleting the folder itself.
        /// </summary>
        /// <param name="ropHardDeleteMessagesAndSubfoldersRequest">RopHardDeleteMessagesAndSubfoldersRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopHardDeleteMessagesAndSubfolders.</param>
        /// <param name="responseSOHTable">Server objects handles in RopHardDeleteMessagesAndSubfoldersResponse.</param>
        /// <returns>RopHardDeleteMessagesAndSubfoldersResponse object.</returns>
        public RopHardDeleteMessagesAndSubfoldersResponse HardDeleteMessagesAndSubfolders(
            RopHardDeleteMessagesAndSubfoldersRequest ropHardDeleteMessagesAndSubfoldersRequest,
            uint insideObjHandle,
            ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropHardDeleteMessagesAndSubfoldersRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopHardDeleteMessagesAndSubfoldersResponse ropHardDeleteMessagesAndSubfoldersResponse = (RopHardDeleteMessagesAndSubfoldersResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropHardDeleteMessagesAndSubfoldersResponse.ReturnValue)
            {
                this.VerifyRopHardDeleteMessagesAndSubfolders(ropHardDeleteMessagesAndSubfoldersResponse);
            }
            #endregion

            return ropHardDeleteMessagesAndSubfoldersResponse;
        }

        /// <summary>
        /// Deletes one or more messages from a folder. 
        /// </summary>
        /// <param name="ropDeleteMessagesRequest">RopDeleteMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopDeleteMessages.</param>
        /// <param name="responseSOHTable">Server objects handles in RopDeleteMessagesResponse.</param>
        /// <returns>RopDeleteMessagesResponse object.</returns>
        public RopDeleteMessagesResponse DeleteMessages(RopDeleteMessagesRequest ropDeleteMessagesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropDeleteMessagesRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopDeleteMessagesResponse ropDeleteMessagesResponse = (RopDeleteMessagesResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropDeleteMessagesResponse.ReturnValue)
            {
                this.VerifyRopDeleteMessages(ropDeleteMessagesResponse);
            }
            #endregion

            return ropDeleteMessagesResponse;
        }

        /// <summary>
        /// Hard deletes one or more messages that are listed in the request buffer. 
        /// </summary>
        /// <param name="ropHardDeleteMessagesRequest">RopHardDeleteMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopHardDeleteMessages.</param>
        /// <param name="responseSOHTable">Server objects handles in RopHardDeleteMessagesResponse.</param>
        /// <returns>RopHardDeleteMessagesResponse object.</returns>
        public RopHardDeleteMessagesResponse HardDeleteMessages(RopHardDeleteMessagesRequest ropHardDeleteMessagesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropHardDeleteMessagesRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopHardDeleteMessagesResponse ropHardDeleteMessagesResponse = (RopHardDeleteMessagesResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropHardDeleteMessagesResponse.ReturnValue)
            {
                this.VerifyRopHardDeleteMessages(ropHardDeleteMessagesResponse);
            }
            #endregion

            return ropHardDeleteMessagesResponse;
        }

        /// <summary>
        /// Retrieves the hierarchy table for a folder. 
        /// </summary>
        /// <param name="ropGetHierarchyTableRequest">RopGetHierarchyTableRequest object.</param>
        /// <param name="insideObjHandle">Server object handle RopGetHierarchyTable.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetHierarchyTableResponse.</param>
        /// <returns>RopGetHierarchyTableResponse object.</returns>
        public RopGetHierarchyTableResponse GetHierarchyTable(RopGetHierarchyTableRequest ropGetHierarchyTableRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropGetHierarchyTableRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            RopGetHierarchyTableResponse ropGetHierarchyTableResponse = (RopGetHierarchyTableResponse)temp;

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropGetHierarchyTableResponse.ReturnValue)
            {
                this.VerifyRopGetHierarchyTable(ropGetHierarchyTableResponse);
            }
            #endregion

            return ropGetHierarchyTableResponse;
        }

        /// <summary>
        /// Retrieves the contents table for a folder.
        /// </summary>
        /// <param name="ropGetContentsTableRequest">RopGetContentsTableRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopGetContentsTable.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetContentsTableResponse.</param>
        /// <returns>RopGetContentsTableResponse object.</returns>
        public RopGetContentsTableResponse GetContentsTable(RopGetContentsTableRequest ropGetContentsTableRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            RopGetContentsTableResponse ropGetContentsTableResponse;
            bool contentsTableLocked = true;
            int count = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            do
            {
                this.ExcuteRopCall((ISerializable)ropGetContentsTableRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
                ropGetContentsTableResponse = (RopGetContentsTableResponse)temp;

                // The contents table was locked by another database operation.
                if (ropGetContentsTableResponse.ReturnValue == 4294965994)
                {
                    Site.Log.Add(LogEntryKind.Comment, " JET_errTableLocked:" + ropGetContentsTableResponse.ReturnValue);
                    Thread.Sleep(waitTime);
                }
                else
                {
                    contentsTableLocked = false;
                }

                if (count > retryCount)
                {
                    break;
                }

                count++;
            }
            while (contentsTableLocked);

            #region Capture Code
            // The ReturnValue equal to 0x00000000 indicate ROP operation success
            if (0x00000000 == ropGetContentsTableResponse.ReturnValue)
            {
                this.VerifyRopGetContentsTable(ropGetContentsTableResponse);
            }
            #endregion

            return ropGetContentsTableResponse;
        }

        /// <summary>
        /// Set folder object properties.
        /// </summary>
        /// <param name="ropSetPropertiesRequest">RopSetPropertiesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in SetProperties.</param>
        /// <param name="responseSOHTable">Server objects handles in RopSetPropertiesResponse.</param>
        /// <returns>RopSetPropertiesResponse object.</returns>
        public RopSetPropertiesResponse SetFolderObjectProperties(RopSetPropertiesRequest ropSetPropertiesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object temp = new object();
            this.ExcuteRopCall((ISerializable)ropSetPropertiesRequest, insideObjHandle, ref temp, ref responseSOHTable, ref this.rawData);
            return (RopSetPropertiesResponse)temp;
        }

        /// <summary>
        /// Get folder object specific properties.
        /// </summary>
        /// <param name="ropGetPropertiesSpecificRequest">RopGetPropertiesSpecificRequest object</param>
        /// <param name="insideObjHandle">Server object handle in GetPropertiesSpecific.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetPropertiesSpecificResponse.</param>
        /// <returns>RopGetPropertiesSpecificResponse object.</returns>
        public RopGetPropertiesSpecificResponse GetFolderObjectSpecificProperties(RopGetPropertiesSpecificRequest ropGetPropertiesSpecificRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable)
        {
            object ropResponse = new object();
            this.ExcuteRopCall((ISerializable)ropGetPropertiesSpecificRequest, insideObjHandle, ref ropResponse, ref responseSOHTable, ref this.rawData);
            RopGetPropertiesSpecificResponse response = (RopGetPropertiesSpecificResponse)ropResponse;

            if (0x00000000 == response.ReturnValue)
            {
                // The getPropertiesSpecificResponse.ReturnValue equals 0 means that this Rop is successful.
                // So the propertyTags in getPropertiesSpecificRequest is correct
                this.VerifyGetFolderSpecificProperties(ropGetPropertiesSpecificRequest.PropertyTags);
            }

            return response;
        }

        /// <summary>
        /// Get all properties of a folder object.
        /// </summary>
        /// <param name="inputHandle">The handle specified the folder RopGetPropertiesAll Rop operation performs on.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetPropertiesSpecificResponse.</param>
        /// <returns>RopGetPropertiesAllResponse object.</returns>
        public RopGetPropertiesAllResponse GetFolderPropertiesAll(uint inputHandle, ref List<List<uint>> responseSOHTable)
        {
            object ropResponse = new object();
            RopGetPropertiesAllRequest request = new RopGetPropertiesAllRequest
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = 0,
                PropertySizeLimit = ushort.MaxValue,
                WantUnicode = 0x01
            };
            this.ExcuteRopCall((ISerializable)request, inputHandle, ref ropResponse, ref responseSOHTable, ref this.rawData);
            RopGetPropertiesAllResponse response = (RopGetPropertiesAllResponse)ropResponse;

            if (0x00000000 == response.ReturnValue)
            {
                this.VerifyGetFolderPropertiesAll(response);
            }

            return response;
        }
        #endregion

        #region Help methods

        /// <summary>
        /// Execute a ROP call.
        /// </summary>
        /// <param name="ropRequest">ROP request objects</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <param name="rawData">The original ROP response payload.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        private uint ExcuteRopCall(ISerializable ropRequest, uint insideObjHandle, ref object response, ref List<List<uint>> responseSOHTable, ref byte[] rawData)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };
            List<uint> requestSOH = new List<uint>
            {
                insideObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responses = new List<IDeserializable>();
            responseSOHTable = new List<List<uint>>();

            // 0x10008 specifies the maximum size of the rgbOut buffer to place in Response.
            uint returnValue = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responses, ref responseSOHTable, ref rawData, 0x10008);

            if (returnValue == OxcRpcErrorCode.ECRpcFormat)
            {
                throw new FormatException("Error RPC Format");
            }

            if (responses != null)
            {
                if (responses.Count > 0)
                {
                    response = responses[0];

                    this.VerifyTransport();
                    this.VerifyRPCLayerRequirement();
                }
            }
            else
            {
                response = null;
            }

            return returnValue;
        }

        /// <summary>
        /// Execute a ROP call.
        /// </summary>
        /// <param name="ropRequest">ROP request objects</param>
        /// <param name="insideObjHandle">Server object handles in request.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <param name="rawData">The original ROP response payload.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        private uint ExcuteRopCall(ISerializable ropRequest, List<uint> insideObjHandle, ref object response, ref List<List<uint>> responseSOHTable, ref byte[] rawData)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };
            List<IDeserializable> responses = new List<IDeserializable>();
            responseSOHTable = new List<List<uint>>();

            // 0x10008 specifies the maximum size of the rgbOut buffer to place in Response.
            uint returnValue = this.oxcropsClient.RopCall(requestRops, insideObjHandle, ref responses, ref responseSOHTable, ref rawData, 0x10008);
            if (returnValue == OxcRpcErrorCode.ECRpcFormat)
            {
                throw new FormatException("Error RPC Format");
            }

            if (responses != null)
            {
                if (responses.Count > 0)
                {
                    response = responses[0];

                    this.VerifyTransport();
                    this.VerifyRPCLayerRequirement();
                }
            }
            else
            {
                response = null;
            }

            return returnValue;
        }
        #endregion
    }
}