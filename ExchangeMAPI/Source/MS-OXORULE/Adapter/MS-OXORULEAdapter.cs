namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter of MS-OXORULE.
    /// </summary>
    public partial class MS_OXORULEAdapter : ManagedAdapterBase, IMS_OXORULEAdapter
    {
        #region Variables

        /// <summary>
        /// Indicate whether RPC connected.
        /// </summary>
        private bool isConnected = false;

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// The NSPIAdapter instance.
        /// </summary>
        private NSPIAdapter nspiAdapter;
        /// <summary>
        /// String server name.
        /// </summary>
        private string server;

        /// <summary>
        /// String user1 ESSDN.
        /// </summary>
        private string user1ESSDN;

        /// <summary>
        /// String user2 ESSDN.
        /// </summary>
        private string user2ESSDN;

        /// <summary>
        /// String domain name.
        /// </summary>
        private string domain;
   
        /// <summary>
        /// Mailbox guid value.
        /// </summary>
        private byte[] mailboxGUID;

        /// <summary>
        /// RPC raw data.
        /// </summary>
        private byte[] rawData;

        /// <summary>
        /// Response of ROP operation.
        /// </summary>
        private object response;

        /// <summary>
        /// Server objects handles.
        /// </summary>
        private List<List<uint>> responseSOHs;

        /// <summary>
        /// This enum is used to specify the ROP operation is performed on  ExtendedRule, DAM, DEM or StandardRule.
        /// </summary>
        private TargetOfRop targetOfRop;

        /// <summary>
        /// Gets or sets indicate the Rop operation is performed for DAM, DEM or ExtendedRules.
        /// </summary>
        public TargetOfRop TargetOfRop
        {
            get { return this.targetOfRop; }
            set { this.targetOfRop = value; }
        }
        #endregion

        #region Defined in MS-OXORULE
        /// <summary>
        /// This ROP gets the rules table of a folder.
        /// </summary>
        /// <param name="objHandle">This index refers to the location in the Server object handle table used to find the handle for this operation.</param>
        /// <param name="tableFlags">These Flags control the Type of table. The possible values are specified in [MS-OXORULE].</param>
        /// <param name="getRulesTableResponse">Structure of RopGetRulesTableResponse.</param>
        /// <returns>Table handle.</returns>
        public uint RopGetRulesTable(uint objHandle, TableFlags tableFlags, out RopGetRulesTableResponse getRulesTableResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopGetRulesTableRequest getRulesTableRequest;

            getRulesTableRequest.RopId = 0x3F;
            getRulesTableRequest.LogonId = 0x00;
            getRulesTableRequest.InputHandleIndex = 0x00;
            getRulesTableRequest.OutputHandleIndex = 0x01;
            getRulesTableRequest.TableFlags = (byte)tableFlags;

            this.responseSOHs = this.DoRPCCall(getRulesTableRequest, objHandle, ref this.response, ref this.rawData);
            getRulesTableResponse = (RopGetRulesTableResponse)this.response;
            uint tableHandle = this.responseSOHs[0][getRulesTableResponse.OutputHandleIndex];

            // Verify the response of RopGetRulesTable 
            this.VerifyRopGetRulesTable(getRulesTableResponse, getRulesTableRequest);
            return tableHandle;
        }

        /// <summary>
        /// This ROP updates the entry IDs in the deferred action messages.
        /// </summary>
        /// <param name="objHandle">This index refers to the location in the Server object handle table used to find the handle for this operation.</param>
        /// <param name="serverEntryId">This value specifies the ID of the message on the server.</param>
        /// <param name="clientEntryId">This value specifies the ID of the downloaded message on the client.</param>
        /// <returns>Structure of RopUpdateDeferredActionMessagesResponse.</returns>
        public RopUpdateDeferredActionMessagesResponse RopUpdateDeferredActionMessages(uint objHandle, byte[] serverEntryId, byte[] clientEntryId)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopUpdateDeferredActionMessagesRequest updateDeferredActionMessagesRequest;
            RopUpdateDeferredActionMessagesResponse updateDeferredActionMessagesResponse;

            updateDeferredActionMessagesRequest.RopId = 0x57;
            updateDeferredActionMessagesRequest.LogonId = 0x00;
            updateDeferredActionMessagesRequest.InputHandleIndex = 0x00;
            updateDeferredActionMessagesRequest.ServerEntryIdSize = 0;
            if (serverEntryId != null)
            {
                updateDeferredActionMessagesRequest.ServerEntryIdSize = (ushort)serverEntryId.Length;
            }

            updateDeferredActionMessagesRequest.ServerEntryId = serverEntryId;
            updateDeferredActionMessagesRequest.ClientEntryIdSize = 0;
            if (clientEntryId != null)
            {
                updateDeferredActionMessagesRequest.ClientEntryIdSize = (ushort)clientEntryId.Length;
            }

            updateDeferredActionMessagesRequest.ClientEntryId = clientEntryId;

            this.responseSOHs = this.DoRPCCall(updateDeferredActionMessagesRequest, objHandle, ref this.response, ref this.rawData);
            updateDeferredActionMessagesResponse = (RopUpdateDeferredActionMessagesResponse)this.response;

            return updateDeferredActionMessagesResponse;
        }

        /// <summary>
        /// This ROP modifies the rules associated with a folder.
        /// </summary>
        /// <param name="objHandle">This index refers to the handle in the Server object handle table used as input for this operation.</param>
        /// <param name="modifyRulesFlags">The possible values are specified in [MS-OXORULE]. These Flags specify behavior of this operation.</param>
        /// <param name="ruleData">An array of RuleData structures, each of which specifies details about a standard rule.</param>
        /// <returns>Structure of RopModifyRulesResponse.</returns>
        public RopModifyRulesResponse RopModifyRules(uint objHandle, ModifyRuleFlag modifyRulesFlags, RuleData[] ruleData)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopModifyRulesRequest modifyRulesRequest;
            RopModifyRulesResponse modifyRulesResponse;

            modifyRulesRequest.RopId = 0x41;
            modifyRulesRequest.LogonId = 0x0;
            modifyRulesRequest.InputHandleIndex = 0x00;
            modifyRulesRequest.ModifyRulesFlags = (byte)modifyRulesFlags;
            modifyRulesRequest.RulesCount = 0;
            if (ruleData != null)
            {
                modifyRulesRequest.RulesCount = (ushort)ruleData.Length;
            }

            modifyRulesRequest.RulesData = ruleData;

            this.responseSOHs = this.DoRPCCall(modifyRulesRequest, objHandle, ref this.response, ref this.rawData);
            modifyRulesResponse = (RopModifyRulesResponse)this.response;
            
            // Verify the response of RopModifyRules
            this.VerifyRopModifyRules(modifyRulesResponse, modifyRulesRequest);

            // modifyRulesResponse.ReturnValue equals 0 means that RopModifyRules is successful.
            // So the ruleDatas in modifyRulesRequest is correct.
            if (modifyRulesResponse.ReturnValue == 0)
            {
                this.VerifyPropertiesInRuleData(ruleData[0]);
            }

            return modifyRulesResponse;
        }
        #endregion

        /// <summary>
        /// Initialize test environment
        /// </summary>
        /// <param name="testSite">Test Site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXORULE";
            Common.MergeConfiguration(this.Site);
            this.server = Common.GetConfigurationPropertyValue(Constants.Server, this.Site);
            this.user1ESSDN = Common.GetConfigurationPropertyValue(Constants.User1ESSDN, this.Site) + "\0";
            this.domain = Common.GetConfigurationPropertyValue(Constants.Domain, this.Site);
            this.user2ESSDN = Common.GetConfigurationPropertyValue(Constants.User2ESSDN, this.Site) + "\0";
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
            this.nspiAdapter = new NSPIAdapter(this.Site);
        }

        /// <summary>
        /// Release RPC binding and destroy RPC context handle.
        /// </summary>
        public void CleanUp()
        {
            this.oxcropsClient.Disconnect();
        }

        /// <summary>
        /// Create RPC connection.
        /// </summary>
        /// <param name="connectionType">The type of connection.</param>
        /// <param name="userName">A string value indicates the domain account name that connects to server.</param>
        /// <param name="userESSDN">A string that identifies user who is making the EcDoConnectEx call</param>
        /// <param name="userPassword">A string value indicates the password of the user which is used.</param>
        /// <returns>Identify if the connection has been established.</returns>
        public bool Connect(ConnectionType connectionType, string userName, string userESSDN, string userPassword)
        {
            if (this.isConnected == true)
            {
                bool isDisconnect = this.oxcropsClient.Disconnect();
                Site.Assert.IsTrue(isDisconnect, "The RPC connection should be disconnected!");
            }

            this.isConnected = this.oxcropsClient.Connect(this.server, connectionType, userESSDN, this.domain, userName, userPassword);
            return this.isConnected;
        }

        /// <summary>
        /// Get properties from NSPI table.
        /// </summary>
        /// <param name="server">Server address.</param>
        /// <param name="userName">The value of user name.</param>
        /// <param name="domain">The value of Domain.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="columns">PropertyTags to be query.</param>
        /// <returns>Results in PropertyRowSet format.</returns>
        public PropertyRowSet_r? GetRecipientInfo(string server, string userName, string domain, string password, PropertyTagArray_r? columns)
        {
            #region Call NspiBind method to initiate a session between the client and the server.
            uint flags = 0;
            STAT stat = new STAT();
            stat.CodePage = 0x4e4; // Set a valid code page.
            stat.TemplateLocale = 0x409; // Set a valid LCID.
            stat.SortLocale = 0x409; // Set a valid LCID.
                                     // Set value for serverGuid
            FlatUID_r guid = new FlatUID_r
            {
                Ab = new byte[16]
            };
            FlatUID_r? serverGuid = guid;


            ErrorCodeValue result = this.nspiAdapter.NspiBind(flags, stat, ref serverGuid);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result, "NspiBind should return Success!");
            #endregion

            #region Call NspiQueryRows method to get the recipient information.
            stat.ContainerID = 0; // Set the container id to the id of default global address book container
            uint tableCount = 0;
            uint[] table = null;
            uint requestCount = 5000;
            PropertyRowSet_r? propertyRowSet = null;

            result = this.nspiAdapter.NspiQueryRows(flags, ref stat, tableCount, table, requestCount, columns, out propertyRowSet);
            Site.Assert.AreEqual<ErrorCodeValue>(ErrorCodeValue.Success, result, "NspiQueryRows should return Success!");
            #endregion

            uint returnValue = this.nspiAdapter.NspiUnbind(0);
            Site.Assert.AreEqual<uint>(1, returnValue, "NspiUnbind method should return 1 (Success).");
            return propertyRowSet;
        }

        #region Related ROPs
        /// <summary>
        /// This ROP gets specific properties of a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">This field specifies the properties requested.</param>
        /// <returns>Structure of RopGetPropertiesSpecificResponse.</returns>
        public RopGetPropertiesSpecificResponse RopGetPropertiesSpecific(uint objHandle, PropertyTag[] propertyTags)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest;
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;

            getPropertiesSpecificRequest.RopId = 0x07;
            getPropertiesSpecificRequest.LogonId = 0x00;
            getPropertiesSpecificRequest.InputHandleIndex = 0x00;
            getPropertiesSpecificRequest.PropertySizeLimit = 0x00;
            getPropertiesSpecificRequest.WantUnicode = 0x01;
            if (propertyTags != null)
            {
                getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                getPropertiesSpecificRequest.PropertyTagCount = 0x00;
            }
            
            getPropertiesSpecificRequest.PropertyTags = propertyTags;
            this.responseSOHs = this.DoRPCCall(getPropertiesSpecificRequest, objHandle, ref this.response, ref this.rawData);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;

            // The getPropertiesSpecificResponse.ReturnValue equals 0 means that this ROP is successful.
            // So the propertyTags in getPropertiesSpecificRequest is correct
            if (getPropertiesSpecificResponse.ReturnValue == 0)
            {
                this.VerifyPropertiesSpecific(propertyTags);
            }

            return getPropertiesSpecificResponse;
        }

        /// <summary>
        /// This ROP adds or modifies recipients on a message. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="recipientColumns">Array of PropertyTag structures. The number of structures contained in this field is specified by the ColumnCount field. This field specifies the property values that can be included for each recipient.</param>
        /// <param name="recipientRows">List of ModifyRecipientRow structures. The number of structures contained in this field is specified by the RowCount field. .</param>
        /// <returns>Response of RopModifyRecipients.</returns>
        public RopModifyRecipientsResponse RopModifyRecipients(uint objHandle, PropertyTag[] recipientColumns, ModifyRecipientRow[] recipientRows)
        {
            RopModifyRecipientsRequest modifyRecipientsRequest;
            RopModifyRecipientsResponse modifyRecipientsResponse;
            modifyRecipientsRequest.RopId = 0x0E;
            modifyRecipientsRequest.LogonId = 0x00;
            modifyRecipientsRequest.InputHandleIndex = 0x00;

            modifyRecipientsRequest.ColumnCount = (ushort)recipientColumns.Length;
            modifyRecipientsRequest.RecipientColumns = recipientColumns;
            modifyRecipientsRequest.RowCount = (ushort)recipientRows.Length;
            modifyRecipientsRequest.RecipientRows = recipientRows;
            this.responseSOHs = this.DoRPCCall(modifyRecipientsRequest, objHandle, ref this.response, ref this.rawData);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)this.response;

            return modifyRecipientsResponse;
        }

        /// <summary>
        /// This ROP submits a message for sending. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="submitFlags">8-bit Flags structure. These Flags specify special behavior for submitting the message.</param>
        /// <returns>Structure of RopSubmitMessageResponse.</returns>
        public RopSubmitMessageResponse RopSubmitMessage(uint objHandle, SubmitFlag submitFlags)
        {
            RopSubmitMessageRequest submitMessageRequest;
            RopSubmitMessageResponse submitMessageResponse;
            submitMessageRequest.RopId = 0x32;
            submitMessageRequest.LogonId = 0x00;
            submitMessageRequest.InputHandleIndex = 0x00;
            submitMessageRequest.SubmitFlags = (byte)submitFlags;
            this.responseSOHs = this.DoRPCCall(submitMessageRequest, objHandle, ref this.response, ref this.rawData);
            submitMessageResponse = (RopSubmitMessageResponse)this.response;
            return submitMessageResponse;
        }

        /// <summary>
        /// This ROP deletes specific properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">Array of PropertyTag structures, this field specifies the property values to be deleted from the object.</param>
        /// <returns>Structure of RopDeletePropertiesResponse.</returns>
        public RopDeletePropertiesResponse RopDeleteProperties(uint objHandle, PropertyTag[] propertyTags)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopDeletePropertiesRequest deletePropertiesRequest;
            RopDeletePropertiesResponse deletePropertiesResponse;

            deletePropertiesRequest.RopId = 0x0B;
            deletePropertiesRequest.LogonId = 0x0;
            deletePropertiesRequest.InputHandleIndex = 0x00;
            if (propertyTags != null)
            {
                deletePropertiesRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                deletePropertiesRequest.PropertyTagCount = 0x00;
            }

            deletePropertiesRequest.PropertyTags = propertyTags;

            this.responseSOHs = this.DoRPCCall(deletePropertiesRequest, objHandle, ref this.response, ref this.rawData);
            deletePropertiesResponse = (RopDeletePropertiesResponse)this.response;
            return deletePropertiesResponse;
        }

        /// <summary>
        /// This ROP gets the contents table of a container. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="tableFlags">8-bit Flags structure. These Flags control the Type of table.</param>+
        /// <param name="tableHandle">Handle of contents table.</param>
        /// <returns>Response of RopGetContentsTable.</returns>
        public RopGetContentsTableResponse RopGetContentsTable(uint handle, ContentTableFlag tableFlags, out uint tableHandle)
        {
            RopGetContentsTableRequest getContentsTableRequest;
            getContentsTableRequest.RopId = 0x05;
            getContentsTableRequest.LogonId = 0x00;
            getContentsTableRequest.InputHandleIndex = 0x00;
            getContentsTableRequest.OutputHandleIndex = 0x01;
            getContentsTableRequest.TableFlags = (byte)tableFlags;
            this.responseSOHs = this.DoRPCCall(getContentsTableRequest, handle, ref this.response, ref this.rawData);
            RopGetContentsTableResponse getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            tableHandle = this.responseSOHs[0][getContentsTableResponse.OutputHandleIndex];

            return getContentsTableResponse;
        }

        /// <summary>
        /// This ROP creates a new subfolder. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="displayName">This value specifies the name of the created folder. .</param>
        /// <param name="comment">This value specifies the folder comment that is associated with the created folder.</param>
        /// <param name="createFolderResponse">Response of this ROP.</param>
        /// <returns>Handle of new folder.</returns>
        public uint RopCreateFolder(uint handle, string displayName, string comment, out RopCreateFolderResponse createFolderResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopCreateFolderRequest createFolderRequest;

            createFolderRequest.RopId = 0x1C;
            createFolderRequest.LogonId = 0x0;
            createFolderRequest.InputHandleIndex = 0x0;
            createFolderRequest.OutputHandleIndex = 0x01;

            // Generic folder
            createFolderRequest.FolderType = 0x01;

            // FALSE
            createFolderRequest.UseUnicodeStrings = 0x0;

            // non-zero(TRUE)
            createFolderRequest.OpenExisting = 0xFF;

            // FALSE
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(displayName + "\0");
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(comment + "\0");
            this.responseSOHs = this.DoRPCCall(createFolderRequest, handle, ref this.response, ref this.rawData);
            createFolderResponse = (RopCreateFolderResponse)this.response;
            uint folderHandle = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];

            return folderHandle;
        }

        /// <summary>
        /// This ROP sets specific properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="taggedPropertyValueArray">Array of TaggedPropertyValue structures, this field specifies the property values to be set on the object.</param>
        /// <returns>Structure of RopSetPropertiesResponse.</returns>
        public RopSetPropertiesResponse RopSetProperties(uint objHandle, TaggedPropertyValue[] taggedPropertyValueArray)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopSetPropertiesRequest setPropertiesRequest;
            RopSetPropertiesResponse setPropertiesResponse;

            setPropertiesRequest.RopId = 0x0A;
            setPropertiesRequest.LogonId = 0x0;
            setPropertiesRequest.InputHandleIndex = 0x00;
            if (taggedPropertyValueArray != null)
            {
                setPropertiesRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;
                ushort count = 2;
                foreach (TaggedPropertyValue tagValue in taggedPropertyValueArray)
                {
                    count += (ushort)(tagValue.Value.Length + 4);
                }

                setPropertiesRequest.PropertyValueSize = count;
            }
            else
            {
                setPropertiesRequest.PropertyValueCount = 0x00;
                setPropertiesRequest.PropertyValueSize = 2;
            }

            setPropertiesRequest.PropertyValues = taggedPropertyValueArray;
            this.responseSOHs = this.DoRPCCall(setPropertiesRequest, objHandle, ref this.response, ref this.rawData);
            setPropertiesResponse = (RopSetPropertiesResponse)this.response;

            return setPropertiesResponse;
        }

        /// <summary>
        /// This ROP commits the changes made to a message. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <returns>Response of this ROP.</returns>
        public RopSaveChangesMessageResponse RopSaveChangesMessage(uint handle)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            saveChangesMessageRequest.RopId = 0x0C;
            saveChangesMessageRequest.LogonId = 0x0;
            saveChangesMessageRequest.InputHandleIndex = 0x0;
            saveChangesMessageRequest.ResponseHandleIndex = 0x01;
            saveChangesMessageRequest.SaveFlags = 0x0C;
            this.responseSOHs = this.DoRPCCall(saveChangesMessageRequest, handle, ref this.response, ref this.rawData);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

            return saveChangesMessageResponse;
        }

        /// <summary>
        /// This ROP gets all properties on a message.
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertySizeLimit">This value specifies the maximum Size allowed for a property value returned.</param>
        /// <param name="wantUnicode">This value specifies whether to return string properties in Unicode.</param>
        /// <returns>Structure of RopGetPropertiesAllResponse.</returns>
        public RopGetPropertiesAllResponse RopGetPropertiesAll(uint objHandle, ushort propertySizeLimit, ushort wantUnicode)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopGetPropertiesAllRequest getPropertiesAllRequest;
            RopGetPropertiesAllResponse getPropertiesAllResponse;

            getPropertiesAllRequest.RopId = 0x08;
            getPropertiesAllRequest.LogonId = 0x0;
            getPropertiesAllRequest.InputHandleIndex = 0x00;
            getPropertiesAllRequest.PropertySizeLimit = propertySizeLimit;
            getPropertiesAllRequest.WantUnicode = wantUnicode;

            this.responseSOHs = this.DoRPCCall(getPropertiesAllRequest, objHandle, ref this.response, ref this.rawData);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;

            // The targetOfRop field indicates that where the Properties get from.
            switch (this.targetOfRop)
            {
                // Verify ExtendedRule's Properties
                case TargetOfRop.ForExtendedRules:
                    this.VerifyRopGetPropertiesAllForExtendedRules(getPropertiesAllResponse);
                    break;

                // Verify DEM's Properties
                case TargetOfRop.ForDEM:
                    this.VerifyRopGetPropertiesAllForDEM(getPropertiesAllResponse);
                    break;
            }

            return getPropertiesAllResponse;
        }

        /// <summary>
        /// This ROP creates a Message object in a mailbox. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">This value identifies the parent folder.</param>
        /// <param name="isFAIMessage">8-bit Boolean. This value specifies whether the message is a folder associated information (FAI) message.</param>
        /// <param name="createMessageResponse">Response of this ROP.</param>
        /// <returns>Handle of the create message.</returns>
        public uint RopCreateMessage(uint handle, ulong folderId, byte isFAIMessage, out RopCreateMessageResponse createMessageResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopCreateMessageRequest req;
            req.RopId = 0x06;
            req.LogonId = 0;
            req.InputHandleIndex = 0;
            req.OutputHandleIndex = 1;
            req.CodePageId = 0x0fff;
            req.FolderId = folderId;
            req.AssociatedFlag = isFAIMessage;

            this.responseSOHs = this.DoRPCCall(req, handle, ref this.response, ref this.rawData);
            createMessageResponse = (RopCreateMessageResponse)this.response;

            return this.responseSOHs[0][createMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP opens an existing message in a mailbox. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">64-bit identifier. This value identifies the parent folder of the message to be opened.</param>
        /// <param name="messageId">64-bit identifier. This value identifies the message to be opened.</param>
        /// <param name="openMessageResponse">Response of this ROP.</param>
        /// <returns>Handle of the open message.</returns>
        public uint RopOpenMessage(uint handle, ulong folderId, ulong messageId, out RopOpenMessageResponse openMessageResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopOpenMessageRequest req;
            req.RopId = 0x03;
            req.LogonId = 0;
            req.InputHandleIndex = 0;
            req.OutputHandleIndex = 1;

            // Use the same codepage as logon object used
            req.CodePageId = 0x0fff;
            req.FolderId = folderId;

            // Read and write
            req.OpenModeFlags = 0x01;
            req.MessageId = messageId;

            this.responseSOHs = this.DoRPCCall(req, handle, ref this.response, ref this.rawData);
            openMessageResponse = (RopOpenMessageResponse)this.response;

            return this.responseSOHs[0][openMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP logs on to a mailbox or public folder. 
        /// </summary>
        /// <param name="logonType">This Type specifies ongoing action on the mailbox or public folder.</param>
        /// <param name="userESSDN">A string that identifies the user to log on to the server.</param>
        /// <param name="logonResponse">Response of this ROP.</param>
        /// <returns>Handle of logon object.</returns>
        public uint RopLogon(LogonType logonType, string userESSDN, out RopLogonResponse logonResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopLogonRequest logonRequest;
            uint objHandle1 = 0;

            byte[] rawData1 = null;
            object response1 = null;
            List<List<uint>> responseSOHs1 = null;
            logonRequest.RopId = 0xFE;
            logonRequest.LogonId = 0x0;
            logonRequest.OutputHandleIndex = 0x0;

            logonRequest.StoreState = 0;

            if (LogonType.PublicFolder == logonType)
            {
                // Logon to public folders
                logonRequest.LogonFlags = 0x00;

                // Logon to public folders
                logonRequest.OpenFlags = 0x01000002;
                logonRequest.EssdnSize = 0;
                logonRequest.Essdn = null;
            }
            else
            {
                // Logon to a private mailbox
                logonRequest.LogonFlags = 0x01;

                // USE_PER_MDB_REPLID_MAPPING flag.
                logonRequest.OpenFlags = 0x01000000;
                if (userESSDN + "\0" == this.user1ESSDN)
                {
                    logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(this.user1ESSDN);
                    logonRequest.Essdn = Encoding.ASCII.GetBytes(this.user1ESSDN);
                }
                else
                {
                    logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(this.user2ESSDN);
                    logonRequest.Essdn = Encoding.ASCII.GetBytes(this.user2ESSDN);
                }
            }

            responseSOHs1 = this.DoRPCCall(logonRequest, objHandle1, ref response1, ref rawData1);
            logonResponse = (RopLogonResponse)response1;
            this.mailboxGUID = logonResponse.MailboxGuid;
            uint handle = responseSOHs1[0][logonResponse.OutputHandleIndex];
            return handle;
        }

        /// <summary>
        /// This ROP opens an existing folder in a mailbox.  
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="folderId">64-bit identifier. This identifier specifies the folder to be opened.</param>
        /// <param name="openFolderResponse">Response of this ROP.</param>
        /// <returns>Handle of the open folder.</returns>
        public uint RopOpenFolder(uint handle, ulong folderId, out RopOpenFolderResponse openFolderResponse)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopOpenFolderRequest openFolderRequest;
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;

            // Open Inbox here
            openFolderRequest.FolderId = folderId;

            // Opening an existing folder
            openFolderRequest.OpenModeFlags = 0x0;
            this.responseSOHs = this.DoRPCCall(openFolderRequest, handle, ref this.response, ref this.rawData);
            openFolderResponse = (RopOpenFolderResponse)this.response;
            uint openedFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            return openedFolderHandle;
        }

        /// <summary>
        /// This ROP deletes all messages and subfolders from a folder. 
        /// </summary>
        /// <param name="handle">Handle to operate.</param>
        /// <param name="wantAsynchronous">8-bit Boolean. This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.</param>
        /// <returns>Response of this ROP.</returns>
        public RopEmptyFolderResponse RopEmptyFolder(uint handle, byte wantAsynchronous)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopEmptyFolderRequest emptyFolderRequest;
            RopEmptyFolderResponse emptyFolderResponse;
            emptyFolderRequest.RopId = 0x58;
            emptyFolderRequest.LogonId = 0x00;
            emptyFolderRequest.InputHandleIndex = 0x0;

            // Synchronously
            emptyFolderRequest.WantAsynchronous = wantAsynchronous;

            // TRUE: delete all messages
            emptyFolderRequest.WantDeleteAssociated = 0xFF;
            this.responseSOHs = this.DoRPCCall(emptyFolderRequest, handle, ref this.response, ref this.rawData);
            emptyFolderResponse = (RopEmptyFolderResponse)this.response;

            return emptyFolderResponse;
        }

        /// <summary>
        /// This ROP sets the properties visible on a table. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="setColumnsFlags">8-bit Flags structure. These Flags control this operation.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <returns>Response of this ROP.</returns>
        public RopSetColumnsResponse RopSetColumns(uint objHandle, byte setColumnsFlags, PropertyTag[] propertyTags)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopSetColumnsRequest setColumnsRequest;
            RopSetColumnsResponse setColumnsResponse;

            setColumnsRequest.RopId = 0x12;
            setColumnsRequest.LogonId = 0x00;
            setColumnsRequest.InputHandleIndex = 0x00;
            setColumnsRequest.SetColumnsFlags = setColumnsFlags;
            setColumnsRequest.PropertyTagCount = 0;
            if (propertyTags != null)
            {
                setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }

            setColumnsRequest.PropertyTags = propertyTags;
            this.responseSOHs = this.DoRPCCall(setColumnsRequest, objHandle, ref this.response, ref this.rawData);
            setColumnsResponse = (RopSetColumnsResponse)this.response;

            return setColumnsResponse;
        }

        /// <summary>
        /// This ROP retrieves rows from a table. 
        /// </summary>
        /// <param name="objHandle">Handle to operate.</param>
        /// <param name="queryRowsFlags">8-bit Flags structure. The possible values are specified in [MS-OXCTABL]. These Flags control this operation.</param>
        /// <param name="forwardRead">8-bit Boolean. This value specifies the direction to read rows.</param>
        /// <param name="rowCount">Unsigned 16-bit integer. This value specifies the number of requested rows.</param>
        /// <returns>Response of this ROP.</returns>
        public RopQueryRowsResponse RopQueryRows(uint objHandle, byte queryRowsFlags, byte forwardRead, ushort rowCount)
        {
            this.rawData = null;
            this.response = null;
            this.responseSOHs = null;

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = queryRowsFlags;
            queryRowsRequest.ForwardRead = forwardRead;
            queryRowsRequest.RowCount = rowCount;

            this.responseSOHs = this.DoRPCCall(queryRowsRequest, objHandle, ref this.response, ref this.rawData);
            queryRowsResponse = (RopQueryRowsResponse)this.response;

            return queryRowsResponse;
        }

        /// <summary>
        /// Get LongTermId from object id.
        /// </summary>
        /// <param name="objHandle">object handle.</param>
        /// <param name="objId">Object id value.</param>
        /// <returns>ROP response.</returns>
        public RopLongTermIdFromIdResponse GetLongTermId(uint objHandle, ulong objId)
        {
            RopLongTermIdFromIdRequest longTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse longTermIdFromIdResponse = new RopLongTermIdFromIdResponse();

            longTermIdFromIdRequest.InputHandleIndex = 0x00;
            longTermIdFromIdRequest.LogonId = 0x00;
            longTermIdFromIdRequest.RopId = 0x43;
            longTermIdFromIdRequest.ObjectId = objId;
            this.responseSOHs = this.DoRPCCall(longTermIdFromIdRequest, objHandle, ref this.response, ref this.rawData);
            longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.response;
            return longTermIdFromIdResponse;
        }

        /// <summary>
        /// Delete specific folder.
        /// </summary>
        /// <param name="objHandle">object handle .</param>
        /// <param name="folderId">ID of the folder will be deleted.</param>
        /// <returns>ROP response of RopDeleteFolder.</returns>
        public RopDeleteFolderResponse RopDeleteFolder(uint objHandle, ulong folderId)
        {
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest();
            RopDeleteFolderResponse deleteFolderResponse = new RopDeleteFolderResponse();
            deleteFolderRequest.RopId = 0x1D;
            deleteFolderRequest.LogonId = 0x00;
            deleteFolderRequest.InputHandleIndex = 0x00;

            // Delete all messages and subfolders
            deleteFolderRequest.DeleteFolderFlags = 0x15;

            // Folder to be deleted
            deleteFolderRequest.FolderId = folderId;
            this.responseSOHs = this.DoRPCCall(deleteFolderRequest, objHandle, ref this.response, ref this.rawData);
            deleteFolderResponse = (RopDeleteFolderResponse)this.response;
            return deleteFolderResponse;
        }

        /// <summary>
        /// Release resources.
        /// </summary>
        /// <param name="handle">Unsigned integer value indicates the Server object Handle</param>
        public void ReleaseRop(uint handle)
        {
            RopReleaseRequest relR = new RopReleaseRequest
            {
                InputHandleIndex = 0x00,
                RopId = (byte)RopId.RopRelease,
                LogonId = 0x00
            };

            this.responseSOHs = this.DoRPCCall(relR, handle, ref this.response, ref this.rawData);
        }
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
        /// <returns>Return the value of ReplyTemplateGUID.</returns>
        public byte[] CreateReplyTemplate(uint inboxFolderHandle, ulong inboxFolderID, bool isOOFReplyTemplate, string templateSubject, TaggedPropertyValue[] addedProperties, out ulong messageId, out uint messageHandler)
        {
            // Create a new FAI message in the inbox folder
            RopCreateMessageResponse ropCreateMessageResponse = new RopCreateMessageResponse();
            messageHandler = this.RopCreateMessage(inboxFolderHandle, inboxFolderID, Convert.ToByte(true), out ropCreateMessageResponse);

            #region Set the new created message's properties
            TaggedPropertyValue[] replyTemplateProperties = new TaggedPropertyValue[3];

            // PidTagMessageClass
            replyTemplateProperties[0] = new TaggedPropertyValue();
            PropertyTag replyTemplatePropertiesPropertyTag = new PropertyTag
            {
                PropertyType = (ushort)PropertyType.PtypString,
                PropertyId = (ushort)PropertyId.PidTagMessageClass
            };
            replyTemplateProperties[0].PropertyTag = replyTemplatePropertiesPropertyTag;
            if (isOOFReplyTemplate == true)
            {
                replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.OOFReplyTemplate + "\0");
            }
            else
            {
                replyTemplateProperties[0].Value = Encoding.Unicode.GetBytes(Constants.ReplyTemplate + "\0");
            }

            // PidTagReplyTemplateId
            replyTemplateProperties[1] = new TaggedPropertyValue();
            PropertyTag pidTagReplyTemplateIdPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagReplyTemplateId,
                PropertyType = (ushort)PropertyType.PtypBinary
            };
            replyTemplateProperties[1].PropertyTag = pidTagReplyTemplateIdPropertyTag;
            Guid newGuid = System.Guid.NewGuid();
            replyTemplateProperties[1].Value = Common.AddInt16LengthBeforeBinaryArray(newGuid.ToByteArray());

            // PidTagSubject
            replyTemplateProperties[2] = new TaggedPropertyValue();
            PropertyTag pidTagSubjectPropertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            replyTemplateProperties[2].PropertyTag = pidTagSubjectPropertyTag;
            replyTemplateProperties[2].Value = Encoding.Unicode.GetBytes(templateSubject + "\0");

            this.RopSetProperties(messageHandler, replyTemplateProperties);
            this.RopSetProperties(messageHandler, addedProperties);
            #endregion

            // Save changes of the message
            RopSaveChangesMessageResponse ropSaveChangesMessagResponse = this.RopSaveChangesMessage(messageHandler);
            messageId = ropSaveChangesMessagResponse.MessageId;
            return newGuid.ToByteArray();
        }
        
        /// <summary>
        /// Query properties in contents table.
        /// </summary>
        /// <param name="tableHandle">Handle of a specific contents table.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <returns>Response of this query rows.</returns>
        public RopQueryRowsResponse QueryPropertiesInTable(uint tableHandle, PropertyTag[] propertyTags)
        {
            // Set the properties in propertyTags to be visible.
            RopSetColumnsResponse setColumnsResponse = this.RopSetColumns(tableHandle, 0x00, propertyTags);

            // Query properties values specified in propertyTags.
            RopQueryRowsResponse queryRowsResponse = this.RopQueryRows(tableHandle, 0x00, 0x01, 1000);

            // That the two Rops are successful means that the propertyTags in request is correct
            if (setColumnsResponse.ReturnValue == 0 && queryRowsResponse.ReturnValue == 0 && queryRowsResponse.RowData.PropertyRows != null)
            {
                // Verify PropertyTags
                this.VerifyPropertiesInTable(propertyTags, queryRowsResponse);

                for (int i = 0; i < propertyTags.Length; i++)
                {
                    // If the property queried is PidTagRuleActions
                    if (propertyTags[i].PropertyId == (ushort)PropertyId.PidTagRuleActions)
                    {
                        for (int j = 0; j < queryRowsResponse.RowData.PropertyRows.Count; j++)
                        {
                            // Verify structure RuleAction 
                            RuleAction ruleAction = new RuleAction();
                            ruleAction.Deserialize(queryRowsResponse.RowData.PropertyRows[j].PropertyValues[i].Value);
                            this.VerifyRuleAction(ruleAction);
                        }
                    }

                    // If the property queried is PidTagExtendedRuleMessageActions
                    if (propertyTags[i].PropertyId == (ushort)PropertyId.PidTagExtendedRuleMessageActions)
                    {
                        for (int j = 0; j < queryRowsResponse.RowData.PropertyRows.Count; j++)
                        {
                            if (BitConverter.ToUInt32(queryRowsResponse.RowData.PropertyRows[j].PropertyValues[i].Value, 0) != (uint)ErrorCodeValue.NotFound)
                            {
                                // Verify structure RuleAction 
                                ExtendedRuleActions ruleAction = new ExtendedRuleActions();
                                byte[] extendedRuleMessageActionBuffer = new byte[queryRowsResponse.RowData.PropertyRows[j].PropertyValues[i].Value.Length - 2];
                                Array.Copy(queryRowsResponse.RowData.PropertyRows[j].PropertyValues[i].Value, 2, extendedRuleMessageActionBuffer, 0, queryRowsResponse.RowData.PropertyRows[j].PropertyValues[i].Value.Length - 2);
                                ruleAction.Deserialize(extendedRuleMessageActionBuffer);
                                this.VerifyExtendRuleAction(ruleAction.RuleActionBuffer);
                            }
                        }
                    }
                }
            }

            return queryRowsResponse;
        }

        /// <summary>
        /// Get notification detail from server.
        /// </summary>
        /// <returns>Notify ROP response.</returns>
        public RopNotifyResponse NotificationProcess()
        {
            List<IDeserializable> responseRops = new List<IDeserializable>();
            RopNotifyResponse notifyResponse = new RopNotifyResponse();
            uint ret = this.oxcropsClient.RopCall(null, null, ref responseRops, ref this.responseSOHs, ref this.rawData, 0x10008);
            Site.Assert.AreEqual<uint>(OxcRpcErrorCode.ECNone, ret, "ROP call should be successful here, the error code is: {0}", ret);

            foreach (IDeserializable response in responseRops)
            {
                string responseName = response.GetType().Name;
                if (responseName == Constants.NameOfRopNotifyResponse)
                {
                    notifyResponse = (RopNotifyResponse)response;
                    break;
                }
            }

            return notifyResponse;
        }

        /// <summary>
        /// Get folder EntryID bytes array.
        /// </summary>
        /// <param name="storeObjectType">Identify the store object is a mailbox or a public folder.</param>
        /// <param name="objHandle">Logon handle.</param>
        /// <param name="folderid">Folder id value.</param>
        /// <returns>Folder EntryID bytes array.</returns>
        public byte[] GetFolderEntryId(StoreObjectType storeObjectType, uint objHandle, ulong folderid)
        {
            // Get folder longterm id.
            RopLongTermIdFromIdResponse longTermIdFromId = this.GetLongTermId(objHandle, folderid);
            FolderEntryID folderEntryId;
            if (storeObjectType == StoreObjectType.Mailbox)
            {
                folderEntryId = new FolderEntryID(storeObjectType, this.mailboxGUID, longTermIdFromId.LongTermId.DatabaseGuid, longTermIdFromId.LongTermId.GlobalCounter);
            }
            else
            {
                byte[] providerUID = new byte[] { 0x1A, 0x44, 0x73, 0x90, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A };
                folderEntryId = new FolderEntryID(storeObjectType, providerUID, longTermIdFromId.LongTermId.DatabaseGuid, longTermIdFromId.LongTermId.GlobalCounter);
            }

            this.VerifyFolderEntryID(folderEntryId, storeObjectType);
            return folderEntryId.Serialize();
        }

        /// <summary>
        /// Get message EntryID bytes array.
        /// </summary>
        /// <param name="folderHandle">Folder handle which the message exist.</param>
        /// <param name="folderId">Folder id value.</param>
        /// <param name="messageHandle">message handle.</param>
        /// <param name="messageId">Message id value.</param>
        /// <returns>Message EntryID bytes array.</returns>
        public byte[] GetMessageEntryId(uint folderHandle, ulong folderId, uint messageHandle, ulong messageId)
        {
            // Get the message longterm ID.
            RopLongTermIdFromIdResponse ropLongTermIdOfMessage = this.GetLongTermId(messageHandle, messageId);

            // Get inbox folder's longterm ID.
            RopLongTermIdFromIdResponse ropLongTermIdOfInboxFolder = this.GetLongTermId(folderHandle, folderId);

            MessageEntryID messageEntryId;

            // Get message's entry ID.
            messageEntryId = new MessageEntryID(this.mailboxGUID, ropLongTermIdOfInboxFolder.LongTermId.DatabaseGuid, ropLongTermIdOfInboxFolder.LongTermId.GlobalCounter, ropLongTermIdOfMessage.LongTermId.DatabaseGuid, ropLongTermIdOfMessage.LongTermId.GlobalCounter);

            this.VerifyMessageEntryID(messageEntryId);
            return messageEntryId.Serialize();
        }
        #endregion

        #region Private method
        /// <summary>
        /// Process and send ROP request.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="objHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <returns>Server objects handles in response.</returns>
        private List<List<uint>> DoRPCCall(ISerializable ropRequest, uint objHandle, ref object response, ref byte[] rawData)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>
            {
                objHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF.
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            // 0x10008 specifies the maximum size of the rgbOut buffer to place in Response.
            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, 0x10008);
            Site.Assert.AreEqual<uint>(OxcRpcErrorCode.ECNone, ret, "ROP call should be successful here, the error code is: {0}", ret);
            if (responseRops.Count != 0)
            {
                response = responseRops[0];
            }

            this.VerifyMAPITransport();
            return responseSOHs;
        }
        #endregion
    }
}