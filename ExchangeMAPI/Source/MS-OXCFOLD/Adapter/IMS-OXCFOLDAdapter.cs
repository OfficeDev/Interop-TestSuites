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
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-OXCFOLD adapter.
    /// </summary>
    public interface IMS_OXCFOLDAdapter : IAdapter
    {
        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <returns>True if connect successfully; otherwise, false.</returns>
        bool DoConnect(ConnectionType connectionType);

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">the type of connection</param>
        /// <param name="userDN">UserDN used to connect server.</param>
        /// <param name="domain">The domain the server is deployed.</param>
        /// <param name="userName">The domain account name.</param>
        /// <param name="password">Password value.</param>
        /// <returns>True if connect server successfully; otherwise, false.</returns>
        bool DoConnect(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password);

        /// <summary>
        /// Client calls it to disconnect the connection with Server.
        /// </summary>
        /// <returns>True if  disconnect server successfully; otherwise, false.</returns>
        bool DoDisconnect();

        /// <summary>
        /// Sends ROP request with single operation and single input object handle with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="ropResponse">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        uint DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object ropResponse, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Sends ROP request with single operation and multiple input object handles with expected SuccessResponse.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="insideObjHandle">The list of server object handles in request.</param>
        /// <param name="ropResponse">ROP response object.</param>
        /// <param name="responseSOHTable">Server objects handles in response.</param>
        /// <returns>An unsigned integer get from server. 0 indicates success, other values indicate failure.</returns>
        uint DoRopCall(ISerializable ropRequest, List<uint> insideObjHandle, ref object ropResponse, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Creates either public folders or private mailbox folders. 
        /// </summary>
        /// <param name="ropCreateFolderRequest">RopCreateFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopCreateFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopCreateFolderResponse.</param>
        /// <returns>RopCreateFolderResponse object.</returns>
        RopCreateFolderResponse CreateFolder(RopCreateFolderRequest ropCreateFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Opens an existing folder.
        /// </summary>
        /// <param name="ropOpenFolderRequest">RopOpenFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopOpenFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopOpenFolderResponse.</param>
        /// <returns>RopOpenFolderResponse object.</returns>
        RopOpenFolderResponse OpenFolder(RopOpenFolderRequest ropOpenFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Removes a subfolder. 
        /// </summary>
        /// <param name="ropDeleteFolderRequest">RopDeleteFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopDeleteFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopDeleteFolderResponse.</param>
        /// <returns>RopDeleteFolderResponse object.</returns>
        RopDeleteFolderResponse DeleteFolder(RopDeleteFolderRequest ropDeleteFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Establishes search criteria for a search folder. 
        /// </summary>
        /// <param name="ropSetSearchCriteriaRequest">RopSetSearchCriteriaRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopSetSearchCriteriaRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopSetSearchCriteriaResponse.</param>
        /// <returns>RopSetSearchCriteriaResponse object.</returns>
        RopSetSearchCriteriaResponse SetSearchCriteria(RopSetSearchCriteriaRequest ropSetSearchCriteriaRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Obtains the search criteria and the status of a search for a search folder. 
        /// </summary>
        /// <param name="ropGetSearchCriteriaRequest">RopGetSearchCriteriaRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopGetSearchCriteriaRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetSearchCriteriaResponse.</param>
        /// <returns>RopGetSearchCriteriaResponse object.</returns>
        RopGetSearchCriteriaResponse GetSearchCriteria(RopGetSearchCriteriaRequest ropGetSearchCriteriaRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Moves or copies messages from a source folder to a destination folder. 
        /// </summary>
        /// <param name="ropMoveCopyMessagesRequest">RopMoveCopyMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopMoveCopyMessagesRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopMoveCopyMessagesResponse.</param>
        /// <returns>RopMoveCopyMessagesResponse object.</returns>
        RopMoveCopyMessagesResponse MoveCopyMessages(RopMoveCopyMessagesRequest ropMoveCopyMessagesRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Moves a folder from one parent to another.
        /// </summary>
        /// <param name="ropMoveFolderRequest">RopMoveFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopMoveFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopMoveFolderResponse.</param>
        /// <returns>RopMoveFolderResponse object.</returns>
        RopMoveFolderResponse MoveFolder(RopMoveFolderRequest ropMoveFolderRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Creates a new folder on the destination parent folder, copying the properties and content of the source folder to the new folder.
        /// </summary>
        /// <param name="ropCopyFolderRequest">RopCopyFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handles in RopCopyFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopCopyFolderResponse.</param>
        /// <returns>RopCopyFolderResponse object.</returns>
        RopCopyFolderResponse CopyFolder(RopCopyFolderRequest ropCopyFolderRequest, List<uint> insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Soft deletes all messages and subfolders from a folder without deleting the folder itself. 
        /// </summary>
        /// <param name="ropEmptyFolderRequest">RopEmptyFolderRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in EmptyFolderRequest.</param>
        /// <param name="responseSOHTable">Server objects handles in RopEmptyFolderResponse.</param>
        /// <returns>RopEmptyFolderResponse object.</returns>
        RopEmptyFolderResponse EmptyFolder(RopEmptyFolderRequest ropEmptyFolderRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Hard deletes all messages and subfolders from a folder without deleting the folder itself.
        /// </summary>
        /// <param name="ropHardDeleteMessagesAndSubfoldersRequest">RopHardDeleteMessagesAndSubfoldersRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopHardDeleteMessagesAndSubfolders.</param>
        /// <param name="responseSOHTable">Server objects handles in RopHardDeleteMessagesAndSubfoldersResponse.</param>
        /// <returns>RopHardDeleteMessagesAndSubfoldersResponse object.</returns>
        RopHardDeleteMessagesAndSubfoldersResponse HardDeleteMessagesAndSubfolders(
            RopHardDeleteMessagesAndSubfoldersRequest ropHardDeleteMessagesAndSubfoldersRequest,
            uint insideObjHandle,
            ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Deletes one or more messages from a folder. 
        /// </summary>
        /// <param name="ropDeleteMessagesRequest">RopDeleteMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopDeleteMessages.</param>
        /// <param name="responseSOHTable">Server objects handles in RopDeleteMessagesResponse.</param>
        /// <returns>RopDeleteMessagesResponse object.</returns>
        RopDeleteMessagesResponse DeleteMessages(RopDeleteMessagesRequest ropDeleteMessagesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Hard deletes one or more messages that are listed in the request buffer. 
        /// </summary>
        /// <param name="ropHardDeleteMessagesRequest">RopHardDeleteMessagesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopHardDeleteMessages.</param>
        /// <param name="responseSOHTable">Server objects handles in RopHardDeleteMessagesResponse.</param>
        /// <returns>RopHardDeleteMessagesResponse object.</returns>
        RopHardDeleteMessagesResponse HardDeleteMessages(RopHardDeleteMessagesRequest ropHardDeleteMessagesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Retrieves the hierarchy table for a folder. 
        /// </summary>
        /// <param name="ropGetHierarchyTableRequest">RopGetHierarchyTableRequest object.</param>
        /// <param name="insideObjHandle">Server object handle RopGetHierarchyTable.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetHierarchyTableResponse.</param>
        /// <returns>RopGetHierarchyTableResponse object.</returns>
        RopGetHierarchyTableResponse GetHierarchyTable(RopGetHierarchyTableRequest ropGetHierarchyTableRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Retrieve the contents table for a folder.
        /// </summary>
        /// <param name="ropGetContentsTableRequest">RopGetContentsTableRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in RopGetContentsTable.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetContentsTableResponse.</param>
        /// <returns>RopGetContentsTableResponse object.</returns>
        RopGetContentsTableResponse GetContentsTable(RopGetContentsTableRequest ropGetContentsTableRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Set folder object properties.
        /// </summary>
        /// <param name="ropSetPropertiesRequest">RopSetPropertiesRequest object.</param>
        /// <param name="insideObjHandle">Server object handle in SetProperties.</param>
        /// <param name="responseSOHTable">Server objects handles in RopSetPropertiesResponse.</param>
        /// <returns>RopSetPropertiesResponse object.</returns>
        RopSetPropertiesResponse SetFolderObjectProperties(RopSetPropertiesRequest ropSetPropertiesRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Get folder object specific properties.
        /// </summary>
        /// <param name="ropGetPropertiesSpecificRequest">RopGetPropertiesSpecificRequest object</param>
        /// <param name="insideObjHandle">Server object handle in GetPropertiesSpecific.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetPropertiesSpecificResponse.</param>
        /// <returns>RopGetPropertiesSpecificResponse object.</returns>
        RopGetPropertiesSpecificResponse GetFolderObjectSpecificProperties(RopGetPropertiesSpecificRequest ropGetPropertiesSpecificRequest, uint insideObjHandle, ref List<List<uint>> responseSOHTable);

        /// <summary>
        /// Get all properties of a folder object.
        /// </summary>
        /// <param name="inputHandle">The handle specified the folder RopGetPropertiesAll Rop operation performs on.</param>
        /// <param name="responseSOHTable">Server objects handles in RopGetPropertiesSpecificResponse.</param>
        /// <returns>RopGetPropertiesAllResponse object.</returns>
        RopGetPropertiesAllResponse GetFolderPropertiesAll(uint inputHandle, ref List<List<uint>> responseSOHTable);
    }
}