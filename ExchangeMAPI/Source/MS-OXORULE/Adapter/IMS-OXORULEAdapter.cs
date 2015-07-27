//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Interface of MS-OXORULEAdapter.
    /// </summary>
    public interface IMS_OXORULEAdapter : IAdapter
    {
        /// <summary>
        /// Gets or sets the TargetOfRop.
        /// </summary>
        TargetOfRop TargetOfRop
        {
            get;
            set;
        }

        #region Defined in MS-OXORULE
        /// <summary>
        /// This ROP gets the rules table of a folder.
        /// </summary>
        /// <param name="objHandle">This index refers to the location in the Server object handle table used to find the handle for this operation.</param>
        /// <param name="tableFlags">These Flags control the Type of table. The possible values are specified in [MS-OXORULE].</param>
        /// <param name="getRulesTableResponse">Structure of RopGetRulesTableResponse.</param>
        /// <returns>Table handle.</returns>
        uint RopGetRulesTable(uint objHandle, TableFlags tableFlags, out RopGetRulesTableResponse getRulesTableResponse);

        /// <summary>
        /// This ROP updates the entry IDs in the deferred action messages.
        /// </summary>
        /// <param name="objHandle">This index refers to the location in the Server object handle table used to find the handle for this operation.</param>
        /// <param name="serverEntryId">This value specifies the ID of the message on the server.</param>
        /// <param name="clientEntryId">This value specifies the ID of the downloaded message on the client.</param>
        /// <returns>Structure of RopUpdateDeferredActionMessagesResponse.</returns>
        RopUpdateDeferredActionMessagesResponse RopUpdateDeferredActionMessages(uint objHandle, byte[] serverEntryId, byte[] clientEntryId);

        /// <summary>
        /// This ROP modifies the rules associated with a folder.
        /// </summary>
        /// <param name="objHandle">This index refers to the handle in the Server object handle table used as input for this operation.</param>
        /// <param name="modifyRulesFlags">The possible values are specified in [MS-OXORULE]. These Flags specify behavior of this operation.</param>
        /// <param name="ruleData">An array of RuleData structures, each of which specifies details about a standard rule.</param>
        /// <returns>Structure of RopModifyRulesResponse.</returns>
        RopModifyRulesResponse RopModifyRules(uint objHandle, ModifyRuleFlag modifyRulesFlags, RuleData[] ruleData);
        #endregion

        /// <summary>
        /// Clean environment.
        /// </summary>
        void CleanUp();

        /// <summary>
        /// Get properties from nspi table.
        /// </summary>
        /// <param name="server">Server address.</param>
        /// <param name="userName">The value of User name.</param>
        /// <param name="domain">The value of Domain.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="columns">PropertyTags to be query.</param>
        /// <returns>Results in PropertyRowSet format.</returns>
        PropertyRowSet_r? GetRecipientInfo(string server, string userName, string domain, string password, PropertyTagArray_r? columns);

        #region Related ROPs
        /// <summary>
        /// This ROP gets specific properties of a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">This field specifies the properties requested.</param>
        /// <returns>Structure of RopGetPropertiesSpecificResponse.</returns>
        RopGetPropertiesSpecificResponse RopGetPropertiesSpecific(uint objHandle, PropertyTag[] propertyTags);

        /// <summary>
        /// This ROP adds or modifies recipients on a message. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="recipientColumns">Array of PropertyTag structures. The number of structures contained in this field is specified by the ColumnCount field. This field specifies the property values that can be included for each recipient.</param>
        /// <param name="recipientRows">List of ModifyRecipientRow structures. The number of structures contained in this field is specified by the RowCount field. .</param>
        /// <returns>Response of RopModifyRecipients.</returns>
        RopModifyRecipientsResponse RopModifyRecipients(uint objHandle, PropertyTag[] recipientColumns, ModifyRecipientRow[] recipientRows);

        /// <summary>
        /// This ROP submits a message for sending. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="submitFlags">8-bit Flags structure. These Flags specify special behavior for submitting the message.</param>
        /// <returns>Structure of RopSubmitMessageResponse.</returns>
        RopSubmitMessageResponse RopSubmitMessage(uint objHandle, SubmitFlag submitFlags);

        /// <summary>
        /// This ROP deletes specific properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values to be deleted from the object.</param>
        /// <returns>Structure of RopDeletePropertiesResponse.</returns>
        RopDeletePropertiesResponse RopDeleteProperties(uint objHandle, PropertyTag[] propertyTags);

        /// <summary>
        /// This ROP gets the contents table of a container. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="tableFlags">8-bit Flags structure. These Flags control the Type of table.</param>
        /// <param name="tableHandle">Handle of contents table.</param>
        /// <returns>Response of RopGetContentsTable.</returns>
        RopGetContentsTableResponse RopGetContentsTable(uint handle, ContentTableFlag tableFlags, out uint tableHandle);

        /// <summary>
        /// This ROP creates a new subfolder. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="displayName">This value specifies the name of the created folder. .</param>
        /// <param name="comment">This value specifies the folder comment that is associated with the created folder.</param>
        /// <param name="createFolderResponse">Response of this ROP.</param>
        /// <returns>Handle of new folder.</returns>
        uint RopCreateFolder(uint handle, string displayName, string comment, out RopCreateFolderResponse createFolderResponse);

        /// <summary>
        /// This ROP sets specific properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="taggedPropertyValueArray">Array of TaggedPropertyValue structures. This field specifies the property values to be set on the object.</param>
        /// <returns>Structure of RopSetPropertiesResponse.</returns>
        RopSetPropertiesResponse RopSetProperties(uint objHandle, TaggedPropertyValue[] taggedPropertyValueArray);

        /// <summary>
        /// This ROP commits the changes made to a message. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <returns>Response of this ROP.</returns>
        RopSaveChangesMessageResponse RopSaveChangesMessage(uint handle);

        /// <summary>
        /// This ROP gets all properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertySizeLimit">This value specifies the maximum Size allowed for a property value returned.</param>
        /// <param name="wantUnicode">This value specifies whether to return string properties in Unicode.</param>
        /// <returns>Structure of RopGetPropertiesAllResponse.</returns>
        RopGetPropertiesAllResponse RopGetPropertiesAll(uint objHandle, ushort propertySizeLimit, ushort wantUnicode);

        /// <summary>
        /// This ROP creates a Message object in a mailbox. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">This value identifies the parent folder.</param>
        /// <param name="isFAIMessage">8-bit Boolean. This value specifies whether the message is a folder associated information (FAI) message.</param>
        /// <param name="createMessageResponse">Response of this ROP.</param>
        /// <returns>Handle of the create message.</returns>
        uint RopCreateMessage(uint handle, ulong folderId, byte isFAIMessage, out RopCreateMessageResponse createMessageResponse);

        /// <summary>
        /// This ROP opens an existing message in a mailbox. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">64-bit identifier. This value identifies the parent folder of the message to be opened.</param>
        /// <param name="messageId">64-bit identifier. This value identifies the message to be opened.</param>
        /// <param name="openMessageResponse">Response of this ROP.</param>
        /// <returns>Handle of the open message.</returns>
        uint RopOpenMessage(uint handle, ulong folderId, ulong messageId, out RopOpenMessageResponse openMessageResponse);

        /// <summary>
        /// This ROP logs on to a mailbox or public folder. 
        /// </summary>
        /// <param name="logonType">This Type specifies ongoing action on the mailbox or public folder.</param>
        /// <param name="userESSDN">A string that identifies the user to log on to the server.</param>
        /// <param name="logonResponse">Response of this ROP.</param>
        /// <returns>Handle of logon object.</returns>
        uint RopLogon(LogonType logonType, string userESSDN, out RopLogonResponse logonResponse);

        /// <summary>
        /// This ROP opens an existing folder in a mailbox.  
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">64-bit identifier. This identifier specifies the folder to be opened.</param>
        /// <param name="openFolderResponse">Response of this ROP.</param>
        /// <returns>Handle of the open folder.</returns>
        uint RopOpenFolder(uint handle, ulong folderId, out RopOpenFolderResponse openFolderResponse);

        /// <summary>
        /// This ROP deletes all messages and subfolders from a folder. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="wantAsynchronous">8-bit Boolean. This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.</param>
        /// <returns>Response of this ROP.</returns>
        RopEmptyFolderResponse RopEmptyFolder(uint handle, byte wantAsynchronous);

        /// <summary>
        /// This ROP sets the properties visible on a table. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="setColumnsFlags">8-bit Flags structure. These Flags control this operation.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <returns>Response of this ROP.</returns>
        RopSetColumnsResponse RopSetColumns(uint objHandle, byte setColumnsFlags, PropertyTag[] propertyTags);

        /// <summary>
        /// This ROP retrieves rows from a table. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="queryRowsFlags">8-bit Flags structure. The possible values are specified in [MS-OXCTABL]. These Flags control this operation.</param>
        /// <param name="forwardRead">8-bit Boolean. This value specifies the direction to read rows.</param>
        /// <param name="rowCount">Unsigned 16-bit integer. This value specifies the number of requested rows.</param>
        /// <returns>Response of this ROP.</returns>
        RopQueryRowsResponse RopQueryRows(uint objHandle, byte queryRowsFlags, byte forwardRead, ushort rowCount);

        /// <summary>
        /// Delete specific folder.
        /// </summary>
        /// <param name="objHandle">object handle .</param>
        /// <param name="folderId">ID of the folder will be deleted.</param>
        /// <returns>ROP response of RopDeleteFolder.</returns>
        RopDeleteFolderResponse RopDeleteFolder(uint objHandle, ulong folderId);

        /// <summary>
        /// Release resources.
        /// </summary>
        /// <param name="handle">Unsigned integer value indicates the Server object Handle</param>
        void ReleaseRop(uint handle);

        /// <summary>
        /// Create rpc connection.
        /// </summary>
        /// <param name="connectionType">The type of connection.</param>
        /// <param name="userName">A string value indicates the domain account name that connects to server.</param>
        /// <param name="userESSDN">A string that identifies user who is making the EcDoConnectEx call</param>
        /// <param name="userPassword">A string value indicates the password of the user which is used.</param>
        /// <returns>Identify if the connection has been established.</returns>
        bool Connect(ConnectionType connectionType, string userName, string userESSDN, string userPassword);
        #endregion

        #region Help method
        /// <summary>
        /// Create Reply Template use settings in Util.cs file.
        /// It will be used by create Rule of ActioType: OP_REPLY
        /// </summary>
        /// <param name="inboxFolderHandle">The inbox folder's handle.</param>
        /// <param name="inboxFolderID">The inbox folder's ID.</param>
        /// <param name="isOOFReplyTemplate">Indicate whether the template to be created is a template for OP_REPLY or OP_OOF_REPLY .</param>
        /// <param name="templateSubject">The name of the template.</param>
        /// <param name="addedProperties">The properties that need to add to the reply template.</param>
        /// <param name="messageId">Message id of reply template message.</param>
        /// <param name="messageHandler">The reply message Handler.</param>
        /// <returns>Return the ReplyTemplateGUID.</returns>
        byte[] CreateReplyTemplate(uint inboxFolderHandle, ulong inboxFolderID, bool isOOFReplyTemplate, string templateSubject, TaggedPropertyValue[] addedProperties, out ulong messageId, out uint messageHandler);

        /// <summary>
        /// Query properties in contents table.
        /// </summary>
        /// <param name="tableHandle">Handle of a specific contents table.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <returns>Response of this query rows.</returns>
        RopQueryRowsResponse QueryPropertiesInTable(uint tableHandle, PropertyTag[] propertyTags);

        /// <summary>
        /// Get LongTermId from object id.
        /// </summary>
        /// <param name="objHandle">object handle.</param>
        /// <param name="objId">Object id value.</param>
        /// <returns>ROP response.</returns>
        RopLongTermIdFromIdResponse GetLongTermId(uint objHandle, ulong objId);

        /// <summary>
        /// Get notification detail from server.
        /// </summary>
        /// <returns>Notify ROP response.</returns>
        RopNotifyResponse NotificationProcess();

        /// <summary>
        /// Get folder entryid bytes array.
        /// </summary>
        /// <param name="storeObjectType">Identify the store object is a mailbox or a public folder.</param>
        /// <param name="objHandle">Logon handle.</param>
        /// <param name="folderid">Folder id value.</param>
        /// <returns>Folder entryid bytes array.</returns>
        byte[] GetFolderEntryId(StoreObjectType storeObjectType, uint objHandle, ulong folderid);

        /// <summary>
        /// Get message EntryID bytes array.
        /// </summary>
        /// <param name="folderHandle">Folder handle which the message exist.</param>
        /// <param name="folderId">Folder id value.</param>
        /// <param name="messageHandle">message handle.</param>
        /// <param name="messageId">Message id value.</param>
        /// <returns>Message EntryID bytes array.</returns>
        byte[] GetMessageEntryId(uint folderHandle, ulong folderId, uint messageHandle, ulong messageId);
        #endregion
    }
}