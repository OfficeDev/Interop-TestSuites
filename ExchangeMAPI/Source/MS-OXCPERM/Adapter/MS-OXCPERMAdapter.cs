namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Implementation of the MS-OXCPERM adapter. 
    /// </summary>
    public partial class MS_OXCPERMAdapter : ManagedAdapterBase, IMS_OXCPERMAdapter
    {
        #region Fields

        /// <summary>
        /// Indicate null value
        /// </summary>
        private const ulong UInt64NULL = ulong.MinValue;

        /// <summary>
        /// The 32 bit integer indicates error occurs
        /// </summary>
        private const uint UINT32FAILED = 0xFFFFFFFF;

        /// <summary>
        /// The return value from server when the action is performed successfully.
        /// </summary>
        private const uint UINT32SUCCESS = 0x00000000;

        /// <summary>
        /// Indicate the handle special is invalid
        /// </summary>
        private const uint InvalidHandleValue = 0xFFFFFFFF;

        /// <summary>
        /// Null terminal string
        /// </summary>
        private const string StringNullTerminator = "\0";

        /// <summary>
        /// Const string for the prefix of PidTagRecipientDisplayName.
        /// </summary>
        private const string PrefixOfDisplayName = "DisplayName-";

        /// <summary>
        /// "@" string.
        /// </summary>
        private const string At = "@";

        /// <summary>
        /// A dictionary used to store the properties' Id and related input and output method
        /// </summary>
        private static Dictionary<int, IPropertyInterface> parses;

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// SUT control adapter
        /// </summary>
        private IMS_OXCPERMSUTControlAdapter sutControlAdapter;

        /// <summary>
        /// The ROP response payload
        /// </summary>
        private byte[] rawData;

        /// <summary>
        /// The handle which is retrieved by ROP
        /// </summary>
        private uint inobjHandle = 0;

        /// <summary>
        /// The response of ROP
        /// </summary>
        private object response;

        /// <summary>
        /// The response list which retrieved by ROP
        /// </summary>
        private List<List<uint>> responseSOHs;

        /// <summary>
        /// The permissions list that is not changed
        /// </summary>
        private List<PermissionUserInfo> originalPermissionsList;

        /// <summary>
        /// The current permissions list
        /// </summary>
        private List<PermissionUserInfo> currentPermissionsList;

        /// <summary>
        /// The ROP response is retrieved from server
        /// </summary>
        private RopLogonResponse ropLogonResponse;

        /// <summary>
        /// Specifies the user logons to the server.
        /// </summary>
        private string currentLogonUser = null;

        /// <summary>
        /// The message Ids created by the logon user.
        /// </summary>
        private List<ulong> ownedMessageIds = new List<ulong>();

        /// <summary>
        /// The message Ids not created by the logon user.
        /// </summary>
        private List<ulong> messageIdsCreatedByOther = new List<ulong>();

        /// <summary>
        /// The modified message index
        /// </summary>
        private int ownedMessageIndex = 0;

        /// <summary>
        /// The anonymous user.
        /// </summary>
        private string anonymousUser = "Anonymous";

        /// <summary>
        /// The default user.
        /// </summary>
        private string defaultUser = string.Empty;

        /// <summary>
        /// Gets the server name.
        /// </summary>
        public string ServerName
        {
            get { return Common.GetConfigurationPropertyValue("SutComputerName", this.Site); }
        }

        /// <summary>
        /// Gets the domain name.
        /// </summary>
        public string Domain
        {
            get { return Common.GetConfigurationPropertyValue("Domain", this.Site); }
        }

        /// <summary>
        /// Gets the User2.
        /// </summary>
        public string User2
        {
            get { return Common.GetConfigurationPropertyValue("AdminUserName", this.Site); }
        }

        /// <summary>
        /// Gets the password of the User2.
        /// </summary>
        public string AdminUserPassword
        {
            get { return Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site); }
        }

        /// <summary>
        /// Gets the User2 ESSDN.
        /// </summary>
        public string AdminUserEssdn
        {
            get { return Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site); }
        }

        /// <summary>
        /// Gets the User1.
        /// </summary>
        public string User1
        {
            get { return Common.GetConfigurationPropertyValue("User1Name", this.Site); }
        }

        /// <summary>
        /// Gets the password for the User1.
        /// </summary>
        public string User1Password
        {
            get { return Common.GetConfigurationPropertyValue("User1Password", this.Site); }
        }

        /// <summary>
        /// Gets the User1 ESSDN.
        /// </summary>
        public string User1Essdn
        {
            get { return Common.GetConfigurationPropertyValue("User1Essdn", this.Site); }
        }

        #endregion Fields

        #region Initialize TestSuite

        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCPERM";
            Common.MergeConfiguration(testSite);

            // Initialize OxcropsClient instance for logon user.
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        /// <summary>
        /// Revert the permissions list.
        /// </summary>
        public override void Reset()
        {
            this.Logon(this.User2);

            uint folderHandle = this.GetFolderObjectHandle(FolderTypeEnum.CommonFolderType);
            
            // Clean up the folder
            RopHardDeleteMessagesAndSubfoldersRequest deleteSubfoldersOfInboxRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = 0x92,
                InputHandleIndex = 0x00,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xff
            };
            this.DoRopCall(deleteSubfoldersOfInboxRequest, folderHandle, ref this.response, ref this.rawData);

            RopHardDeleteMessagesAndSubfoldersResponse deleteSubfoldersOfInboxResponse = (RopHardDeleteMessagesAndSubfoldersResponse)this.response;
            Site.Assert.AreEqual<uint>(0, deleteSubfoldersOfInboxResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.messageIdsCreatedByOther.Clear();
            this.ownedMessageIds.Clear();

            // Restore the permissions for the user configured by the "User1Name"
            for (int i = 0; i < this.originalPermissionsList.Count; i++)
            {
                if (this.originalPermissionsList[i].PidTagMemberName == this.User1)
                {
                    PermissionData[] revertPermissionData = new PermissionData[1];
                    PermissionData permission = new PermissionData
                    {
                        PermissionDataFlags = 0x01
                    };

                    PropertyValue[] propertyValues = new PropertyValue[2];
                    propertyValues[0] = this.CreateEntryIdProperty(this.originalPermissionsList[i].PidTagMemberName);
                    propertyValues[1] = this.CreateRightsProperty(this.originalPermissionsList[i].PidTagMemberRights);
                    permission.PropertyValueCount = (ushort)propertyValues.Length;
                    permission.PropertyValues = propertyValues;

                    revertPermissionData[0] = permission;
                    RequestBufferFlags requestBufferFlags = new RequestBufferFlags
                    {
                        IsReplaceRowsFlagSet = false
                    };

                    // Call RopModifyPermissions to revert the permissions list.
                    this.responseSOHs = this.DoRopCall(
                        this.CreateModifyPermissionsRequestBuffer(revertPermissionData, requestBufferFlags),
                        folderHandle,
                        ref this.response,
                        ref this.rawData);

                    RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
                    Site.Log.Add(LogEntryKind.Comment, "The return value for restore the permissions in Reset(): {0:X}", modifyPermissionsResponse.ReturnValue);
                    break;
                }
                else
                {
                    this.RemovePermission(FolderTypeEnum.CommonFolderType, this.User1, new RequestBufferFlags());
                }
            }

            this.ReleaseObject(folderHandle);
            base.Reset();
        }

        #endregion

        #region Interface Implementation

        /// <summary>
        /// Initialize the permission list.
        /// </summary>
        public void InitializePermissionList()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && !Common.IsRequirementEnabled(1184, this.Site))
            {
                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. Only Exchange 2013 SP1 supports MAPIHTTP transport.");
            }

            this.currentLogonUser = string.Empty;
            this.sutControlAdapter = Site.GetAdapter<IMS_OXCPERMSUTControlAdapter>();

            parses = new Dictionary<int, IPropertyInterface>();
            this.RegisterToParses();
            uint responseValue = 0;

            // The user configured by "AdminUserName" logons to his own mailbox to retrieve or modify the permissions.
            this.Logon(this.User2);

            // Retrieve the permissions list to restore.
            RequestBufferFlags bufferFlags = new RequestBufferFlags
            {
                IsIncludeFreeBusyFlagSet = true
            };

            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;
            this.rawData = null;
            uint status;

            uint folderHandle = this.GetFolderObjectHandle(folderType);

            status = this.RetrievePermissionsList(folderType, folderHandle, bufferFlags);
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, status, "0 indicates querying permission data operates successfully.");

            bool completeToParse = this.ParseUserListFromRawData(out this.originalPermissionsList, ref this.rawData);
            Site.Assert.IsTrue(completeToParse, "True indicates the permissions list is parsed successfully.");

            this.currentPermissionsList = this.originalPermissionsList;

            responseValue = this.RemovePermission(folderType, this.User1, bufferFlags);
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, responseValue, "0 indicates the server removes permission successfully.");
        }

        /// <summary>
        /// Read the folder's PidTagSecurityDescriptorAsXml property. For more details see section 3.1.4.1 Retrieving Folder Permissions of [MS-OXCPERM]
        /// </summary>
        /// <param name="folderType">The folder type specifies the PidTagSecurityDescriptorAsXml property of the folder is read.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        public uint ReadSecurityDescriptorProperty(FolderTypeEnum folderType)
        {
            uint folderHandle = this.GetFolderObjectHandle(folderType);
            this.DoRopCall(this.CreateOpenStreamRequestBuffer(), (uint)folderHandle, ref this.response, ref this.rawData);
            RopOpenStreamResponse openStreamResponse = (RopOpenStreamResponse)this.response;

            // Release the stream object
            this.ReleaseObject(folderHandle);

            return openStreamResponse.ReturnValue;
        }

        /// <summary>
        /// Get the permission list for a user on the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The user whose permission is returned. The user name is the alias of the user without domain</param>
        /// <param name="requestBufferFlags">The TableFlags or the ModifyFlags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns> 
        public uint GetPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, out List<PermissionTypeEnum> permissionList)
        {
            if (permissionUserName == null)
            {
                throw new ArgumentNullException("permissionUserName");
            }

            permissionList = new List<PermissionTypeEnum>();
            this.rawData = null;

            uint folderHandle = this.GetFolderObjectHandle(folderType);
            uint status = this.RetrievePermissionsList(folderType, folderHandle, requestBufferFlags);
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, status, "0 indicates querying permission data operates successfully.");

            this.VerifyReturnValueForGetPermission();

            bool succeedToParse = this.ParseUserListFromRawData(out this.currentPermissionsList, ref this.rawData);
            if (succeedToParse)
            {
                this.VerifyMessageSyntax();
                this.VerifyPropertiesOfDataStructure(succeedToParse, this.currentPermissionsList);
                this.VerifyProperties(succeedToParse, this.currentPermissionsList);
            }

            uint responseValue = this.ParsePermissionByUserNameFromUserInfoList(permissionUserName, out permissionList, ref this.currentPermissionsList);

            return responseValue;
        }

        /// <summary>
        /// Add a permission for a user to the permission list of the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The user whose permission is added. The user name is the alias of the user without domain</param>
        /// <param name="requestBufferFlags">Request buffer flags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>      
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        public uint AddPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, List<PermissionTypeEnum> permissionList)
        {
            if (permissionUserName == null)
            {
                throw new ArgumentNullException("permissionUserName");
            }

            uint folderHandle = this.GetFolderObjectHandle(folderType);
            uint pidTagMemberRights = this.ParsePermissionFromPermissionTypeListToByteFormat(permissionList);

            PermissionData[] permissionsDataArray = this.SetPermissionDataArrayForAdd(permissionUserName, pidTagMemberRights);

            this.responseSOHs = this.DoRopCall(
                this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, requestBufferFlags),
                folderHandle,
                ref this.response,
                ref this.rawData);
            RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
            this.VerifyModifyPermissionsResponse();
            uint responseValue = modifyPermissionsResponse.ReturnValue;

            this.VerifyReturnValueForModifyPermission();

            if (responseValue == UINT32SUCCESS)
            {
                this.VerifyReturnValueSuccessForModifyPermission(responseValue);
            }

            // RopRelease folder handle 
            this.ReleaseObject(folderHandle);

            return responseValue;
        }

        /// <summary>
        /// Modify the permission list for a user on the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The user whose permission is modified. The user name is the alias of the user without domain</param>
        /// <param name="requestBufferFlags">Request buffer flags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>       
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        public uint ModifyPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, List<PermissionTypeEnum> permissionList)
        {
            if (permissionUserName == null)
            {
                throw new ArgumentNullException("permissionUserName");
            }

            uint folderHandle = this.GetFolderObjectHandle(folderType);

            List<PermissionUserInfo> permissionUserList;
            uint responseValue = UINT32FAILED;
            this.QueryPermission(folderType, out folderHandle, out permissionUserList);

            bool memberIdIsExisted = false;
            ulong pidTagMemberId = this.GetPidTagMemberId(permissionUserList, permissionUserName, out memberIdIsExisted);
            if (memberIdIsExisted)
            {
                uint nrights = this.ParsePermissionFromPermissionTypeListToByteFormat(permissionList);
                PermissionData[] permissionsDataArray = this.SetPermissionDataArrayForModify(pidTagMemberId, nrights);

                this.DoRopCall(this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, requestBufferFlags), folderHandle, ref this.response, ref this.rawData);

                RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
                this.VerifyModifyPermissionsResponse();
                responseValue = modifyPermissionsResponse.ReturnValue;

                this.VerifyReturnValueForModifyPermission();

                if (responseValue == UINT32SUCCESS)
                {
                    this.VerifyReturnValueSuccessForModifyPermission(responseValue);
                }
            }
            else
            {
                Site.Assert.Fail("No pidTagMemberId for the user: {0}", permissionUserName);
            }

            // RopRelease folder handle 
            this.ReleaseObject(folderHandle);

            return responseValue;
        }

        /// <summary>
        ///  Remove a permission for a user from the permission list on the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The user is removed from the permissions list. The user name is the alias of the user without domain</param>
        /// <param name="requestBufferFlags">Request buffer flags</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        public uint RemovePermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags)
        {
            if (permissionUserName == null)
            {
                throw new ArgumentNullException("permissionUserName");
            }

            uint folderHandle = this.GetFolderObjectHandle(folderType);

            List<PermissionUserInfo> permissionUserList;
            uint responseValue = UINT32FAILED;
            this.QueryPermission(folderType, out folderHandle, out permissionUserList);

            bool memberIdIsExisted = false;
            ulong pidTagMemberId = this.GetPidTagMemberId(permissionUserList, permissionUserName, out memberIdIsExisted);
            if (memberIdIsExisted)
            {
                PermissionData[] permissionsDataArray = this.SetPermissionDataArrayForRemove(pidTagMemberId);
                this.DoRopCall(this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, requestBufferFlags), folderHandle, ref this.response, ref this.rawData);

                RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
                this.VerifyModifyPermissionsResponse();
                responseValue = modifyPermissionsResponse.ReturnValue;

                this.VerifyReturnValueForModifyPermission();

                if (responseValue == UINT32SUCCESS)
                {
                    this.VerifyReturnValueSuccessForModifyPermission(responseValue);
                }
            }
            else
            {
                return UINT32SUCCESS;
            }

            // RopRelease folder message 
            this.ReleaseObject(folderHandle);

            return responseValue;
        }

        /// <summary>
        /// Check whether the user has the permission to operate the corresponding behavior specified by permission. 
        /// </summary>
        /// <param name="permission">The permission flag specified in PidTagMemberRights</param>
        /// <param name="userName">The user whose permission is specified in permission argument.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        public uint CheckPidTagMemberRightsBehavior(PermissionTypeEnum permission, string userName)
        {
            uint returnValue = UINT32FAILED;
            string freeBusy = string.Empty;
            string password = string.Empty;

            switch (permission)
            {
                case PermissionTypeEnum.Create:
                    returnValue = this.CreateMessageByLogonUser();

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyCreateFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.CreateSubFolder:
                    returnValue = this.CreateSubFolder(userName);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyCreateSubFolderFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.DeleteAny:
                    returnValue = this.DeleteMessage(this.messageIdsCreatedByOther[0]);
                    if (UINT32SUCCESS != returnValue)
                    {
                        return returnValue;
                    }

                    returnValue = this.DeleteMessage(this.ownedMessageIds[0]);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyDeleteAnyFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.DeleteOwned:
                    returnValue = this.DeleteMessage(this.ownedMessageIds[0]);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyDeleteOwnedFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.EditAny:
                    returnValue = this.EditMessage(this.messageIdsCreatedByOther[0]);
                    if (UINT32SUCCESS != returnValue)
                    {
                        return returnValue;
                    }

                    returnValue = this.EditMessage(this.ownedMessageIds[0]);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyEditAnyFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.EditOwned:
                    returnValue = this.EditMessage(this.ownedMessageIds[0]);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyEditOwnedFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.FolderOwner:
                    returnValue = this.ModifyFolderProperty();

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyFolderOwnerFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.FolderVisible:
                    returnValue = this.CheckFolderVisible();

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyFolderVisibleFlagValue();
                    }

                    return returnValue;

                case PermissionTypeEnum.FreeBusyDetailed:
                    password = this.GetUserPassword(userName);
                    freeBusy = this.sutControlAdapter.GetUserFreeBusyStatus(userName, password);
                    if ("1" == freeBusy)
                    {
                        this.VerifyFreeBusyDetailedFlagValue();
                        return UINT32SUCCESS;
                    }
                    else if ("2" == freeBusy)
                    {
                        return 0x80070005; // AccessDenied
                    }
                    else
                    {
                        return UINT32FAILED;
                    }

                case PermissionTypeEnum.FreeBusySimple:
                    password = this.GetUserPassword(userName);
                    freeBusy = this.sutControlAdapter.GetUserFreeBusyStatus(userName, password);
                    if ("0" == freeBusy)
                    {
                        this.VerifyFreeBusySimpleFlagValue();
                        return UINT32SUCCESS;
                    }
                    else if ("2" == freeBusy)
                    {
                        return 0x80070005; // AccessDenied
                    }
                    else
                    {
                        return UINT32FAILED;
                    }

                case PermissionTypeEnum.ReadAny:

                    returnValue = this.ReadMessage(this.ownedMessageIds[0]);
                    if (UINT32SUCCESS != returnValue)
                    {
                        return returnValue;
                    }

                    returnValue = this.ReadMessage(this.messageIdsCreatedByOther[0]);

                    if (UINT32SUCCESS == returnValue)
                    {
                        this.VerifyReadAnyFlagValue();
                    }

                    return returnValue;

                default:
                    return UINT32FAILED;
            }
        }

        /// <summary>
        /// The user connects to the server and logons to the mailbox of the user configured by "AdminUserName" in ptfconfig.
        /// </summary>
        /// <param name="userName">The user to logon to the mailbox of the user configured by "AdminUserName" in ptfconfig</param>
        public void Logon(string userName)
        {
            bool isConnected;

            // If the user is already connected to the server, not connect again.
            // If the user is not connected to the server, connect it.
            if (!string.Equals(this.currentLogonUser, userName, StringComparison.OrdinalIgnoreCase))
            {
                // Disconnect the previous RPC connection.
                this.oxcropsClient.Disconnect();

                string userDN = string.Empty;
                string password = string.Empty;
                if (string.Equals(userName, this.User2, StringComparison.OrdinalIgnoreCase))
                {
                    userDN = this.AdminUserEssdn;
                    password = this.AdminUserPassword;
                }
                else if (string.Equals(userName, this.User1, StringComparison.OrdinalIgnoreCase))
                {
                    userDN = this.User1Essdn;
                    password = this.User1Password;
                }
                else
                {
                    this.currentLogonUser = null;
                    Site.Assert.Fail("The user {0} is not allowed to logon to the server. Check the user.", userName);
                }

                // The user logons to the User2's mailbox.
                isConnected = this.oxcropsClient.Connect(this.ServerName, ConnectionType.PrivateMailboxServer, userDN, this.Domain, userName, password);
                Site.Assert.IsTrue(isConnected, "True indicates connecting to server successfully.");
                this.ropLogonResponse = this.ROPLogon(LogonType.Mailbox, this.AdminUserEssdn, out this.inobjHandle);
                Site.Assert.AreEqual<uint>(UINT32SUCCESS, this.ropLogonResponse.ReturnValue, "0 indicates the user {0} logs on to the {1}'s mailbox successfully.", userName, this.User2);
                this.VerifyCredential(userName, password, this.ropLogonResponse.ReturnValue);
                this.currentLogonUser = userName;
            }
            else
            {
                Site.Log.Add(LogEntryKind.Comment, "The user {0} is already connected to the server.", userName);
            }
        }

        /// <summary>
        /// Create a new message in the mail box folder of the user configured by "AdminUserName" by the logon user.
        /// </summary>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        public uint CreateMessageByLogonUser()
        {
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;
            createMessageRequest.RopId = 0x06; // RopId 0x06 indicates RopCreateMessage
            createMessageRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            createMessageRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            createMessageRequest.OutputHandleIndex = 0x01; // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            createMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            createMessageRequest.FolderId = this.ropLogonResponse.FolderIds[4]; // Create a message in INBOX which root is mailbox 
            createMessageRequest.AssociatedFlag = 0x00; // NOT an FAI message
            this.responseSOHs = this.DoRopCall(createMessageRequest, this.inobjHandle, ref this.response, ref this.rawData);
            createMessageResponse = (RopCreateMessageResponse)this.response;
            if (UINT32SUCCESS != createMessageResponse.ReturnValue)
            {
                return createMessageResponse.ReturnValue;
            }

            uint targetMessageHandle = this.responseSOHs[0][createMessageResponse.OutputHandleIndex];

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = 0x0C; // RopId 0x0C indicates RopSaveChangesMessage
            saveChangesMessageRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            saveChangesMessageRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            saveChangesMessageRequest.ResponseHandleIndex = 0x01; // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            saveChangesMessageRequest.SaveFlags = 0x0C; // ForceSave
            this.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

            if (saveChangesMessageResponse.ReturnValue == UINT32SUCCESS)
            {
                if (this.currentLogonUser == this.User2)
                {
                    this.messageIdsCreatedByOther.Add(saveChangesMessageResponse.MessageId);
                }
                else
                {
                    this.ownedMessageIds.Add(saveChangesMessageResponse.MessageId);
                }
            }

            // RopRelease created message 
            this.ReleaseObject(targetMessageHandle);

            return UINT32SUCCESS;
        }

        /// <summary>
        /// Disconnect the connection with server.
        /// </summary>
        /// <returns>True indicates disconnecting successfully, otherwise false</returns>
        public bool Disconnect()
        {
            bool ret = this.oxcropsClient.Disconnect();
            return ret;
        }

        /// <summary>
        /// Check the error code AccessDenied when calling RopQueryRows ROP.
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        public uint CheckRopQueryRowsErrorCodeAccessDenied(FolderTypeEnum folderType, string permissionUserName, List<PermissionTypeEnum> permissionList)
        {
            // Get the folder handle
            uint folderHandle = this.GetFolderObjectHandle(folderType);

            RequestBufferFlags requestBufferFlags = new RequestBufferFlags();
            uint status = this.RetrievePermissionsList(folderType, folderHandle, requestBufferFlags);
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, status, "0 indicates querying permission data operates successfully.");

            List<PermissionUserInfo> permissionUserList = new List<PermissionUserInfo>();
            this.ParseUserListFromRawData(out permissionUserList, ref this.rawData);

            bool memberIdIsExisted = false;
            ulong pidTagMemberId = this.GetPidTagMemberId(permissionUserList, permissionUserName, out memberIdIsExisted);

            uint nrights = this.ParsePermissionFromPermissionTypeListToByteFormat(permissionList);
            PermissionData[] permissionsDataArray = this.SetPermissionDataArrayForModify(pidTagMemberId, nrights);

            // Check the user doesn't contain FolderVisible
            Site.Assume.IsFalse(permissionList.Contains(PermissionTypeEnum.FolderVisible), "False indicate the user: {0} doesn't contain FolderVisible.", permissionUserName);

            // Modify the permissions
            this.DoRopCall(this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, requestBufferFlags), folderHandle, ref this.response, ref this.rawData);
            RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, modifyPermissionsResponse.ReturnValue, "0 indicates the RopModifyPermissions operates successfully.");

            status = this.RetrievePermissionsList(folderType, folderHandle, requestBufferFlags);
            return status;
        }
        #endregion Interface Implementation

        #region Private Methods

        /// <summary>
        /// Register a property's ID together with a structure
        /// </summary>
        /// <param name="propertyId">A property's ID</param>
        /// <param name="testClass">A interfaceBase structure</param>
        private static void RegisterForParses(int propertyId, IPropertyInterface testClass)
        {
            parses[propertyId] = testClass;
        }

        /// <summary>
        /// Get the calendar folder id.
        /// </summary>
        /// <returns>The calendar folder id.</returns>
        private ulong GetCalendarFolderId()
        {
            // The calendar folder name is specified in [MS-OXOSFLD]
            const string CalendarFolderName = "Calendar";
            ulong folderId = this.ropLogonResponse.FolderIds[0];

            this.responseSOHs = this.DoRopCall(this.CreateOpenFolderRequestBuffer(folderId), (uint)this.inobjHandle, ref this.response, ref this.rawData);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, openFolderResponse.ReturnValue, "0 indicates the RopOpenFolder operates successfully.");

            uint folderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            this.responseSOHs = this.DoRopCall(this.CreateGetHierarchyTableRequestBuffer(), folderHandle, ref this.response, ref this.rawData);
            RopGetHierarchyTableResponse getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, getHierarchyTableResponse.ReturnValue, "0 indicates the RopGetHierarchyTable operates successfully.");

            uint tableHandle = this.responseSOHs[0][getHierarchyTableResponse.OutputHandleIndex];

            PropertyTag[] propertyTags = this.CreateHierarchyTablePropertyTagsForFolder();
            this.DoRopCall(this.CreateSetColumnsRequestBuffer(propertyTags), tableHandle, ref this.response, ref this.rawData);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, setColumnsResponse.ReturnValue, "0 indicates the RopSetColumns operates successfully.");

            this.DoRopCall(this.CreateQueryRowsRequestBuffer(), tableHandle, ref this.response, ref this.rawData);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, queryRowsResponse.ReturnValue, "0 indicates the RopQueryRows operates successfully.");

            List<List<IPropertyInterface>> responseList = this.QueryRowParse(propertyTags, this.rawData);

            folderId = 0;
            for (int i = 0; i < responseList.Count; i++)
            {
                if (responseList[i][1].OutputValue(responseList[i][1]).Equals(CalendarFolderName))
                {
                    folderId = (ulong)responseList[i][0].OutputValue(responseList[i][0]);
                    break;
                }
            }

            // Release table object
            this.ReleaseObject(tableHandle);

            if (0 == folderId)
            {
                Site.Log.Add(LogEntryKind.Comment, "The folder id isn't queried for calendar folder");
            }

            return folderId;
        }

        /// <summary>
        /// Process and send ROP request.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inobjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <returns>Server objects handles in response.</returns>
        private List<List<uint>> DoRopCall(ISerializable ropRequest, uint inobjHandle, ref object response, ref byte[] rawData)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };
            List<uint> requestSOH = new List<uint>
            {
                inobjHandle
            };

            if (ropRequest != null && Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();
            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, 0x10008);
            if (ret == OxcRpcErrorCode.ECRpcFormat)
            {
                this.Site.Assert.Fail("Error RPC Format");
            }

            this.Site.Assert.AreEqual<uint>(0x0, ret, "If the response is success, the return value is 0x0.");
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

            Type[] paraTypes = new Type[1];
            paraTypes[0] = response.GetType();

            this.VerifyMAPITransport();

            return responseSOHs;
        }

        /// <summary>
        /// Logon onto the Exchange server.
        /// </summary>
        /// <param name="logonType">Specify the logon folder type: PublicFolder, MailBox.</param>
        /// <param name="userDN">Identifies the mailbox to logon.</param>
        /// <param name="objHandle">The server object handle.</param>
        /// <returns>The Logon ROP response.</returns>
        private RopLogonResponse ROPLogon(LogonType logonType, string userDN, out uint objHandle)
        {
            RopLogonRequest logonRequest;
            this.inobjHandle = 0;
            logonRequest.RopId = 0xFE;
            logonRequest.LogonId = 0x0;
            logonRequest.OutputHandleIndex = 0x0;
            logonRequest.StoreState = 0;
            string fullUserDNString = userDN + StringNullTerminator;

            if (LogonType.PublicFolder == logonType)
            {
                logonRequest.LogonFlags = 0x00; // Logon to public folders
                logonRequest.OpenFlags = 0x01000002; // Logon to public folders
                logonRequest.EssdnSize = 0;
                logonRequest.Essdn = null;
            }
            else
            {
                logonRequest.LogonFlags = 0x01; // Logon to a private mailbox
                logonRequest.OpenFlags = 0x01000000;

                logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(fullUserDNString);
                logonRequest.Essdn = Encoding.ASCII.GetBytes(fullUserDNString);
            }

            this.responseSOHs = this.DoRopCall(logonRequest, this.inobjHandle, ref this.response, ref this.rawData);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;

            Site.Assert.AreEqual<uint>(UINT32SUCCESS, logonResponse.ReturnValue, "0 indicates RopLogon operates successfully.");
            objHandle = this.responseSOHs[0][logonResponse.OutputHandleIndex];

            return logonResponse;
        }

        /// <summary>
        /// Call RopRelease ROP to release the object.
        /// </summary>
        /// <param name="objectHandle">Specify the object handle to release</param>
        private void ReleaseObject(uint objectHandle)
        {
            object response = new object();
            byte[] rawData = new byte[0];

            RopReleaseRequest releaseRequest;
            releaseRequest.RopId = 0x01; // RopId 0x01 indicates RopRelease
            releaseRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            releaseRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            this.DoRopCall(releaseRequest, objectHandle, ref response, ref rawData);
        }

        /// <summary>
        /// Retrieve the permission list for a user of the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="folderHandle">Folder handle of the folder specified by argument folderType.</param>
        /// <param name="requestBufferFlags">The TableFlags or the ModifyFlags</param>
        /// <returns>Specify whether the folder handle is valid. 0 indicates valid, others indicate error occurs.</returns>
        private uint RetrievePermissionsList(FolderTypeEnum folderType, uint folderHandle, RequestBufferFlags requestBufferFlags)
        {
            this.responseSOHs = this.DoRopCall(this.CreateGetPermissionsTableRequestBuffer(requestBufferFlags), (uint)folderHandle, ref this.response, ref this.rawData);
            RopGetPermissionsTableResponse getPermissionsTableResponse = (RopGetPermissionsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, getPermissionsTableResponse.ReturnValue, "0 indicates the RopGetPermissionsTable operates successfully.");

            this.VerifyGetPermissionHandle(getPermissionsTableResponse.ReturnValue);
            this.VerifyReturnValueSuccessForGetPermission(getPermissionsTableResponse.ReturnValue);

            uint tableHandle = this.responseSOHs[0][getPermissionsTableResponse.OutputHandleIndex];

            PropertyTag[] propertyTags = this.CreateHierarchyTablePropertyTagsForPermissionUser();
            this.DoRopCall(this.CreateSetColumnsRequestBuffer(propertyTags), tableHandle, ref this.response, ref this.rawData);
            RopSetColumnsResponse ropSetColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, ropSetColumnsResponse.ReturnValue, "0 indicates the RopSetColumns operates successfully.");

            this.DoRopCall(this.CreateQueryRowsRequestBuffer(), tableHandle, ref this.response, ref this.rawData);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;

            return queryRowsResponse.ReturnValue;
        }

        /// <summary>
        /// Read a message in the mail box folder according to the message Id.
        /// </summary>
        /// <param name="messageId">The message Id by which the message identified is to read.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint ReadMessage(ulong messageId)
        {
            RopOpenMessageRequest openMessageRequest;
            RopOpenMessageResponse openMessageResponse;
            openMessageRequest.RopId = 0x03; // RopId 0x03 indicates RopOpenMessage
            openMessageRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            openMessageRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            openMessageRequest.OutputHandleIndex = 0x01; // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            openMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            openMessageRequest.FolderId = this.ropLogonResponse.FolderIds[4]; // Open the message in INBOX folder in which message is created.
            openMessageRequest.OpenModeFlags = 0x00; // The message will be opened as read-only
            openMessageRequest.MessageId = messageId; // Open the saved message
            this.DoRopCall(openMessageRequest, this.inobjHandle, ref this.response, ref this.rawData);
            openMessageResponse = (RopOpenMessageResponse)this.response;
            return openMessageResponse.ReturnValue;
        }

        /// <summary>
        /// Delete the message created by the owner of the mailbox or other user in Inbox.
        /// </summary>
        /// <param name="messageId">The message Id to delete.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint DeleteMessage(ulong messageId)
        {
            // Open the Inbox folder.
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = 0x02,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderId = this.ropLogonResponse.FolderIds[4],
                OpenModeFlags = 0x00
            };
            this.responseSOHs = this.DoRopCall(openFolderRequest, this.inobjHandle, ref this.response, ref this.rawData);
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            if (UINT32SUCCESS != openFolderResponse.ReturnValue)
            {
                return openFolderResponse.ReturnValue;
            }

            uint inboxFolderHandle = this.responseSOHs[0][((RopOpenFolderResponse)this.response).OutputHandleIndex];

            ulong[] messageIdsArrayToDelete = new ulong[1];

            // Delete the message created.
            messageIdsArrayToDelete[0] = messageId;
            RopDeleteMessagesRequest deleteMessagesRequest;
            RopDeleteMessagesResponse deleteMessagesResponse;
            deleteMessagesRequest.RopId = 0x1E;
            deleteMessagesRequest.LogonId = 0x00;
            deleteMessagesRequest.InputHandleIndex = 0x00;
            deleteMessagesRequest.WantAsynchronous = 0x00; // Synchronously
            deleteMessagesRequest.NotifyNonRead = 0x00; // The server does not generate a non-read receipt for the deleted messages
            deleteMessagesRequest.MessageIdCount = (ushort)messageIdsArrayToDelete.Length;
            deleteMessagesRequest.MessageIds = messageIdsArrayToDelete;
            this.DoRopCall(deleteMessagesRequest, inboxFolderHandle, ref this.response, ref this.rawData);
            deleteMessagesResponse = (RopDeleteMessagesResponse)this.response;

            if (deleteMessagesResponse.ReturnValue == UINT32SUCCESS && deleteMessagesResponse.PartialCompletion == 0)
            {
                if (this.ownedMessageIds.Contains(messageId))
                {
                    this.ownedMessageIds.Remove(messageId);
                }
                else
                {
                    this.messageIdsCreatedByOther.Remove(messageId);
                }

                return UINT32SUCCESS;
            }
            else
            {
                if (deleteMessagesResponse.ReturnValue != 0x00000000)
                {
                    return deleteMessagesResponse.ReturnValue;
                }
                else
                {
                    Site.Log.Add(LogEntryKind.Debug, "PartialCompletion of the response to RopDeleteMessages is not 0");
                    return UINT32FAILED;
                }
            }
        }

        /// <summary>
        /// Edit the message created by the owner of the mailbox or other user in Inbox.
        /// </summary>
        /// <param name="messageId">The message Id to edit.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint EditMessage(ulong messageId)
        {
            RopOpenMessageRequest openMessageRequest;
            RopOpenMessageResponse openMessageResponse;
            openMessageRequest.RopId = 0x03; // RopId 0x03 indicates RopOpenMessage
            openMessageRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            openMessageRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            openMessageRequest.OutputHandleIndex = 0x01; // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            openMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            openMessageRequest.FolderId = this.ropLogonResponse.FolderIds[4]; // Open the message in INBOX folder in which message is created.
            openMessageRequest.OpenModeFlags = 0x01; // The message will be opened as read-write
            openMessageRequest.MessageId = messageId; // Open the saved message
            this.responseSOHs = this.DoRopCall(openMessageRequest, this.inobjHandle, ref this.response, ref this.rawData);
            openMessageResponse = (RopOpenMessageResponse)this.response;
            if (UINT32SUCCESS != openMessageResponse.ReturnValue)
            {
                return openMessageResponse.ReturnValue;
            }

            uint targetMessageHandle = this.responseSOHs[0][openMessageResponse.OutputHandleIndex];

            TaggedPropertyValue[] tags = new TaggedPropertyValue[2];

            // PidTagSubject
            PropertyTag pt = new PropertyTag
            {
                PropertyId = 0x0037,
                PropertyType = 0x001F
            };
            TaggedPropertyValue tpv0 = new TaggedPropertyValue
            {
                PropertyTag = pt
            };
            if (this.messageIdsCreatedByOther.Contains(messageId))
            {
                tpv0.Value = Encoding.Unicode.GetBytes("EditTheSubjectCreatedByAdminEditedbyUser1" + StringNullTerminator);
            }
            else
            {
                tpv0.Value = Encoding.Unicode.GetBytes("EditTheSubjectCreatedByCommUserEditedbyUser1: " + this.ownedMessageIndex++ + StringNullTerminator);
            }

            tags[0] = tpv0;

            // PidTagNormalizedSubject
            PropertyTag pt0 = new PropertyTag
            {
                PropertyId = 0x0E1D,
                PropertyType = 0x001F
            };
            TaggedPropertyValue tpv1 = new TaggedPropertyValue
            {
                PropertyTag = pt0,
                Value = tpv0.Value
            };
            tags[1] = tpv1;

            #region Construct the request buffer
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest
            {
                RopId = 0x0A,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertyValueSize = (ushort)(tags[0].Size() + tags[1].Size() + 2),
                PropertyValueCount = (ushort)tags.Length,
                PropertyValues = tags
            };

            #endregion

            this.DoRopCall(setPropertiesRequest, targetMessageHandle, ref this.response, ref this.rawData);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            if (UINT32SUCCESS != setPropertiesResponse.ReturnValue)
            {
                return setPropertiesResponse.ReturnValue;
            }

            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = 0x0C; // RopId 0x0C indicates RopSaveChangesMessage 
            saveChangesMessageRequest.LogonId = 0x00; // The logonId 0x00 is associated with this operation.
            saveChangesMessageRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            saveChangesMessageRequest.ResponseHandleIndex = 0x01; // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            saveChangesMessageRequest.SaveFlags = 0x0C; // ForceSave
            this.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
            if (UINT32SUCCESS != saveChangesMessageResponse.ReturnValue)
            {
                return saveChangesMessageResponse.ReturnValue;
            }

            // Release the message 
            this.ReleaseObject(targetMessageHandle);

            return UINT32SUCCESS;
        }

        /// <summary>
        /// Create a sub-folder in the Inbox folder.
        /// </summary>
        /// <param name="user">The user to create the folder</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint CreateSubFolder(string user)
        {
            // Create a folder with the folder display name
            string folderName = Common.GenerateResourceName(Site, "PermissionToCreateFolder");
            uint inboxFolderHandle;
            uint folderHandle;
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            #region Create the folder

            this.Logon(user);

            // Create the folder in the below steps
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x00;
            openFolderRequest.InputHandleIndex = 0x00;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = this.ropLogonResponse.FolderIds[4]; // Open Inbox folder here
            openFolderRequest.OpenModeFlags = 0x04; // Opening an existing folder
            this.responseSOHs = this.DoRopCall(openFolderRequest, this.inobjHandle, ref this.response, ref this.rawData);

            openFolderResponse = (RopOpenFolderResponse)this.response;
            if (UINT32SUCCESS != openFolderResponse.ReturnValue)
            {
                return openFolderResponse.ReturnValue;
            }

            // Get the Inbox folder handle
            inboxFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;
            createFolderRequest.RopId = 0x1C;
            createFolderRequest.LogonId = 0x00;
            createFolderRequest.InputHandleIndex = 0x00;
            createFolderRequest.OutputHandleIndex = 0x01;
            createFolderRequest.FolderType = 0x01; // Generic folder
            createFolderRequest.UseUnicodeStrings = 0x0; // FALSE
            createFolderRequest.OpenExisting = 0x01; // TRUE
            createFolderRequest.Reserved = 0x0; // FALSE
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(folderName + StringNullTerminator);
            createFolderRequest.Comment = createFolderRequest.DisplayName;
            this.responseSOHs = this.DoRopCall(createFolderRequest, inboxFolderHandle, ref this.response, ref this.rawData);
            createFolderResponse = (RopCreateFolderResponse)this.response;
            if (UINT32SUCCESS != createFolderResponse.ReturnValue)
            {
                return createFolderResponse.ReturnValue;
            }

            folderHandle = this.responseSOHs[0][createFolderResponse.OutputHandleIndex];
            ulong folderId = createFolderResponse.FolderId;

            // Open the created folder to check the folder is created successfully.
            openFolderRequest.FolderId = folderId;
            this.DoRopCall(openFolderRequest, folderHandle, ref this.response, ref this.rawData);
            openFolderResponse = (RopOpenFolderResponse)this.response;

            // Release the folder object.
            this.ReleaseObject(inboxFolderHandle);

            return openFolderResponse.ReturnValue;

            #endregion
        }

        /// <summary>
        /// Modify the folder permissions to check the folder permissions can be changed.
        /// </summary>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint ModifyFolderProperty()
        {
            // Modify the folder permissions.
            FolderTypeEnum folder = FolderTypeEnum.CommonFolderType;
            RequestBufferFlags reqestFlag = new RequestBufferFlags
            {
                IsIncludeFreeBusyFlagSet = true
            };
            List<PermissionTypeEnum> permissions = new List<PermissionTypeEnum>();
            uint returnStatus = this.GetPermission(folder, this.defaultUser, reqestFlag, out permissions);
            if (returnStatus != UINT32SUCCESS)
            {
                return returnStatus;
            }

            // Set the permissions as it has.
            returnStatus = this.ModifyPermission(folder, this.defaultUser, reqestFlag, permissions);
            return returnStatus;
        }

        /// <summary>
        /// Check whether the user can open the Inbox folder and see the hierarchy table of Inbox.
        /// </summary>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        private uint CheckFolderVisible()
        {
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x00;
            openFolderRequest.InputHandleIndex = 0x00;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = this.ropLogonResponse.FolderIds[4]; // Open Inbox folder here
            openFolderRequest.OpenModeFlags = 0x04; // Opening an existing folder
            this.responseSOHs = this.DoRopCall(openFolderRequest, this.inobjHandle, ref this.response, ref this.rawData);

            openFolderResponse = (RopOpenFolderResponse)this.response;
            if (UINT32SUCCESS != openFolderResponse.ReturnValue)
            {
                return openFolderResponse.ReturnValue;
            }

            uint inboxFolderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            RopGetHierarchyTableRequest getHierarchyTableRequest;
            RopGetHierarchyTableResponse getHierarchyTableResponse;
            getHierarchyTableRequest.RopId = 0x04;
            getHierarchyTableRequest.LogonId = 0x00;
            getHierarchyTableRequest.InputHandleIndex = 0x00;
            getHierarchyTableRequest.OutputHandleIndex = 0x01;
            getHierarchyTableRequest.TableFlags = 0x04; // Depth
            this.DoRopCall(getHierarchyTableRequest, inboxFolderHandle, ref this.response, ref this.rawData);
            getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.response;

            // Release the Inbox folder object.
            this.ReleaseObject(inboxFolderHandle);

            return getHierarchyTableResponse.ReturnValue;
        }

        /// <summary>
        /// Create request to open stream
        /// </summary>
        /// <returns>A request is used to open stream</returns>
        private RopOpenStreamRequest CreateOpenStreamRequestBuffer()
        {
            RopOpenStreamRequest openStreamRequest = new RopOpenStreamRequest
            {
                RopId = 0x2B,
                LogonId = 0x0,
                InputHandleIndex = 0x00,
                PropertyTag =
                {
                    PropertyId = 0x0E6A,
                    PropertyType = 0x001F
                },
                OpenModeFlags = 0x00
            };

            return openStreamRequest;
        }

        /// <summary>
        /// Get member id by member name
        /// </summary>
        /// <param name="permissionName">The user name</param>
        /// <param name="permissionUserList">User permission list</param> 
        /// <returns>Return special user's member id</returns>
        private ulong GetMemberIdByName(string permissionName, List<PermissionUserInfo> permissionUserList)
        {
            foreach (PermissionUserInfo permissionUser in permissionUserList)
            {
                if (string.Compare(permissionUser.PidTagMemberName, permissionName, true) == 0)
                {
                    return permissionUser.PidTagMemberId;
                }
            }

            return UInt64NULL;
        }

        /// <summary>
        /// Return the MemberRights
        /// </summary>
        /// <param name="permissionUserList">The permission list of all users</param>
        /// <returns>The List collection of the MemberRights</returns>
        private List<byte[]> GetMemberRights(List<PermissionUserInfo> permissionUserList)
        {
            List<byte[]> memberRightsList = new List<byte[]>();
            foreach (PermissionUserInfo permissionUser in permissionUserList)
            {
                memberRightsList.Add(BitConverter.GetBytes(permissionUser.PidTagMemberRights));
            }

            return memberRightsList;
        }

        /// <summary>
        /// Get the PidTagMemberId values for all users
        /// </summary>
        /// <param name="permissionUserList">The user list</param> 
        /// <returns>Return all of user's id</returns>
        private List<byte[]> GetMemberIDs(List<PermissionUserInfo> permissionUserList)
        {
            List<byte[]> memberIDList = new List<byte[]>();
            foreach (PermissionUserInfo permissionUser in permissionUserList)
            {
                memberIDList.Add(BitConverter.GetBytes(permissionUser.PidTagMemberId));
            }

            return memberIDList;
        }

        /// <summary>
        /// Get special user rights
        /// </summary>
        /// <param name="permissionUserName">The user name</param>
        /// <param name="permissionList">Permission user list</param>
        /// <param name="permissionUserList">Permission mask</param>
        /// <returns>Indicate whether this function is success or not</returns>
        private uint ParsePermissionByUserNameFromUserInfoList(string permissionUserName, out List<PermissionTypeEnum> permissionList, ref List<PermissionUserInfo> permissionUserList)
        {
            if (permissionUserName == null)
            {
                throw new ArgumentNullException("permissionUserName");
            }

            uint nrights = uint.MinValue;

            permissionList = new List<PermissionTypeEnum>();

            foreach (PermissionUserInfo permissionUserInfo in permissionUserList)
            {
                if (string.Compare(permissionUserName, permissionUserInfo.PidTagMemberName, true) == 0)
                {
                    nrights = permissionUserInfo.PidTagMemberRights;
                    this.ParseRightsFromByteToPermissionList(out permissionList, nrights);

                    return UINT32SUCCESS;
                }
            }

            return UINT32FAILED;
        }

        /// <summary>
        /// Create buffer to get PermissionsTable
        /// </summary>
        /// <param name="requestBufferFlags">Request buffer for get permissions table</param>
        /// <returns>Request for getting PermissionsTable</returns>
        private RopGetPermissionsTableRequest CreateGetPermissionsTableRequestBuffer(RequestBufferFlags requestBufferFlags)
        {
            RopGetPermissionsTableRequest getPermissionsTableRequest = new RopGetPermissionsTableRequest
            {
                RopId = 0x3E,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                OutputHandleIndex = 0x01,
                TableFlags = requestBufferFlags.BufferFlags
            };

            return getPermissionsTableRequest;
        }

        /// <summary>
        /// Create buffer to get ModifyPermissions
        /// </summary>
        /// <param name="permissionsDataArray">Permission data array is used to set permission</param>
        /// <param name="requestBufferFlags">requestBufferFlags is used to set flag</param>
        /// <returns>A request is used to modify permissions</returns>
        private RopModifyPermissionsRequest CreateModifyPermissionsRequestBuffer(PermissionData[] permissionsDataArray, RequestBufferFlags requestBufferFlags)
        {
            RopModifyPermissionsRequest modifyPermissionsRequest = new RopModifyPermissionsRequest
            {
                RopId = 0x40,
                LogonId = 0x0,
                InputHandleIndex = 0x00, // Attention InputHandleIndex must be set to 0x00 else 0x02 value will get 1206 error
                ModifyFlags = requestBufferFlags.BufferFlags,
                ModifyCount = (ushort)permissionsDataArray.Length,
                PermissionsData = permissionsDataArray
            };
            return modifyPermissionsRequest;
        }

        /// <summary>
        /// Create request to get permissions table
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <returns>A request is used to get permissions table</returns>
        private RopGetPermissionsTableRequest CreateGetPermissionsTableRequestBuffer(FolderTypeEnum folderType)
        {
            RopGetPermissionsTableRequest getPermissionsTableRequest = new RopGetPermissionsTableRequest
            {
                RopId = 0x3E,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                OutputHandleIndex = 0x01
            };

            if (folderType == FolderTypeEnum.CalendarFolderType)
            {
                getPermissionsTableRequest.TableFlags = 0x02; // IncludeFreeBusy
            }

            return getPermissionsTableRequest;
        }

        /// <summary>
        /// Get the permission data.
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="folderHandle">The handle of the folder</param>
        /// <param name="permissionUserList">Permission user list</param>
        private void QueryPermission(FolderTypeEnum folderType, out uint folderHandle, out List<PermissionUserInfo> permissionUserList)
        {
            permissionUserList = new List<PermissionUserInfo>();

            folderHandle = this.GetFolderObjectHandle(folderType);

            this.responseSOHs = this.DoRopCall(this.CreateGetPermissionsTableRequestBuffer(folderType), (uint)folderHandle, ref this.response, ref this.rawData);
            RopGetPermissionsTableResponse getPermissionsTableResponse = (RopGetPermissionsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, getPermissionsTableResponse.ReturnValue, "0 indicates the RopGetPermissionsTable operates successfully.");

            uint tableHandle = this.responseSOHs[0][getPermissionsTableResponse.OutputHandleIndex];
            PropertyTag[] propertyTags = this.CreateHierarchyTablePropertyTagsForPermissionUser();
            this.DoRopCall(this.CreateSetColumnsRequestBuffer(propertyTags), tableHandle, ref this.response, ref this.rawData);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, setColumnsResponse.ReturnValue, "0 indicates the RopSetColumns operates successfully.");

            this.DoRopCall(this.CreateQueryRowsRequestBuffer(), tableHandle, ref this.response, ref this.rawData);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, queryRowsResponse.ReturnValue, "0 indicates the RopQueryRows operates successfully.");

            if (!this.ParseUserListFromRawData(out permissionUserList, ref this.rawData))
            {
                throw new ArgumentException("Failed to parse the row data from permissions list");
            }

            // RopRelease table message 
            this.ReleaseObject(tableHandle);
        }

        /// <summary>
        /// Create the PropertyTag sent to set column for GetHierarchyTable
        /// </summary>
        /// <returns>The propertyTag contains two properties</returns>
        private PropertyTag[] CreateHierarchyTablePropertyTagsForFolder()
        {
            PropertyTag[] propertyTagArray = new PropertyTag[2];

            // PidTagFolderId
            PropertyTag foldfIdProperty;
            foldfIdProperty.PropertyId = 0x6748;
            foldfIdProperty.PropertyType = 0x0014;
            propertyTagArray[0] = foldfIdProperty;

            // PidTagDisplayName
            PropertyTag displayNameProperty;
            displayNameProperty.PropertyId = 0x3001;
            displayNameProperty.PropertyType = 0x001f;
            propertyTagArray[1] = displayNameProperty;

            return propertyTagArray;
        }

        /// <summary>
        /// Create the PropertyTag sent to set column for GetHierarchyTable
        /// </summary>
        /// <returns>The propertyTag contains two properties</returns>
        private PropertyTag[] CreateHierarchyTablePropertyTagsForPermissionUser()
        {
            PropertyTag[] propertyTags = new PropertyTag[4];

            // PidTagMenberId       
            PropertyTag memberIdProperty;
            memberIdProperty.PropertyId = 0x6671;
            memberIdProperty.PropertyType = 0x0014;
            propertyTags[0] = memberIdProperty;

            // PidTagMemberName
            PropertyTag memberNameProperty;
            memberNameProperty.PropertyId = 0x6672;
            memberNameProperty.PropertyType = 0x001F;
            propertyTags[1] = memberNameProperty;

            // PidTagMemberRights
            PropertyTag memberRightsProperty;
            memberRightsProperty.PropertyId = 0x6673;
            memberRightsProperty.PropertyType = 0x0003;
            propertyTags[2] = memberRightsProperty;

            // PidTagEntryId
            PropertyTag entryIdProperty;
            entryIdProperty.PropertyId = 0x0FFF;
            entryIdProperty.PropertyType = 0x0102;
            propertyTags[3] = entryIdProperty;

            return propertyTags;
        }

        /// <summary>
        /// Parse the return rawData from the QueryRow
        /// </summary>
        /// <param name="propertyTags">The propertyTags define returned properties</param>
        /// <param name="rawData">The rawData contains all the properties</param>
        /// <returns>A double List IPropertyInterface Structure saves all the properties</returns>
        private List<List<IPropertyInterface>> QueryRowParse(PropertyTag[] propertyTags, byte[] rawData)
        {
            List<List<IPropertyInterface>> totalList = new List<List<IPropertyInterface>>();
            IPropertyInterface baseList;
            int count = 0x14;
            int times = rawData[0x11];

            while (times != 0)
            {
                List<IPropertyInterface> list1 = new List<IPropertyInterface>();
                int propertyNumber = propertyTags.GetLength(0);
                while (propertyNumber != 0)
                {
                    baseList = parses[propertyTags[(propertyTags.GetLength(0) - propertyNumber)].PropertyId].InputValue(rawData, ref count);
                    list1.Add(baseList);
                    propertyNumber--;
                }

                count++;
                times--;

                totalList.Add(list1);
            }

            return totalList;
        }

        /// <summary>
        /// Get Folder Object Handle
        /// </summary>
        /// <param name="folder">The folder type</param>
        /// <returns>Return the folder object handle</returns>
        private uint GetFolderObjectHandle(FolderTypeEnum folder)
        {
            ulong folderId = 0;
            if (folder == FolderTypeEnum.CalendarFolderType)
            {
                folderId = this.GetCalendarFolderId();
            }
            else if (folder == FolderTypeEnum.CommonFolderType)
            {
                folderId = this.ropLogonResponse.FolderIds[4];
            }

            RopOpenFolderResponse openFolderResponse;
            this.responseSOHs = this.DoRopCall(this.CreateOpenFolderRequestBuffer(folderId), this.inobjHandle, ref this.response, ref this.rawData);
            openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(UINT32SUCCESS, openFolderResponse.ReturnValue, "0 indicates the RopOpenFolder operates successfully.");

            uint folderHandle = this.responseSOHs[0][openFolderResponse.OutputHandleIndex];

            return folderHandle;
        }

        /// <summary>
        /// Create buffer for open folder.
        /// </summary>
        /// <param name="folderId">The id of folder.</param>
        /// <returns>The return indicate a request which is created for open folder</returns>
        private RopOpenFolderRequest CreateOpenFolderRequestBuffer(ulong folderId)
        {
            RopOpenFolderRequest openFolderRequest;

            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = folderId; // Open Inbox here
            openFolderRequest.OpenModeFlags = 0x0; // Opening an existing folder

            return openFolderRequest;
        }

        /// <summary>
        /// Create buffer for get hierarchy table.
        /// </summary>
        /// <returns>The return indicate a request which is created for get hierarchy table</returns>
        private RopGetHierarchyTableRequest CreateGetHierarchyTableRequestBuffer()
        {
            RopGetHierarchyTableRequest getHierarchyTableRequest;

            getHierarchyTableRequest.RopId = 0x04;
            getHierarchyTableRequest.LogonId = 0x00;
            getHierarchyTableRequest.InputHandleIndex = 0x00;
            getHierarchyTableRequest.OutputHandleIndex = 0x01;
            getHierarchyTableRequest.TableFlags = 0x04; // Depth

            return getHierarchyTableRequest;
        }

        /// <summary>
        /// Create buffer for query rows.
        /// </summary>
        /// <returns>The return indicate a request which is created for query rows</returns>
        private RopQueryRowsRequest CreateQueryRowsRequestBuffer()
        {
            RopQueryRowsRequest queryRowsRequest;

            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = 0x00; // Advance
            queryRowsRequest.ForwardRead = 0x01; // TRUE: read the table forwards
            queryRowsRequest.RowCount = 0xFFFF; // Maximum number of rows to be returned

            return queryRowsRequest;
        }

        /// <summary>
        /// Create buffer for set columns.
        /// </summary>
        /// <param name="propertyTags">Property tags is for set columns.</param>
        /// <returns>The return indicate a request which is created for set columns</returns>
        private RopSetColumnsRequest CreateSetColumnsRequestBuffer(PropertyTag[] propertyTags)
        {
            RopSetColumnsRequest requestBuffer;

            requestBuffer.RopId = 0x12;
            requestBuffer.LogonId = 0x00;
            requestBuffer.InputHandleIndex = 0x00;
            requestBuffer.SetColumnsFlags = 0x00; // Synchronously
            requestBuffer.PropertyTagCount = (ushort)propertyTags.Length;
            requestBuffer.PropertyTags = propertyTags;

            return requestBuffer;
        }

        /// <summary>
        /// Parse the values from server 
        /// </summary>
        /// <param name="permissionUserList">Permission user list</param>
        /// <param name="rawData">The values from server</param>
        /// <returns>If the returnValue is not 0x0000000, return false, else return true.</returns>
        private bool ParseUserListFromRawData(out List<PermissionUserInfo> permissionUserList, ref byte[] rawData)
        {
            permissionUserList = new List<PermissionUserInfo>();

            // Don't need to get the RopId(1byte) InputHandleIndex(1byte).
            int index = 0x0c;

            // Get the ReturnValue, ReturnValue 4 byte.     
            uint returnValue = (uint)BitConverter.ToInt32(rawData, index);
            if (returnValue != UINT32SUCCESS)
            {
                return false;
            }

            // Ignore Origin(1 byte)
            index += 5;

            // RowCout 
            int rowCount = BitConverter.ToInt16(rawData, index);
            index += 2;
            while (rowCount > 0)
            {
                // Flag 1 byte
                ++index;
                PermissionUserInfo permissionUserInfo = new PermissionUserInfo
                {
                    PidTagMemberId = (ulong)BitConverter.ToInt64(rawData, index)
                };

                // PT_I8  PidTagMemberId               
                index += 8;

                // PT_UNICODE  PidTagMemberName
                int countString = index;
                for (; countString < rawData.Length;)
                {
                    if (rawData[countString] == 0x00 && rawData[countString + 1] == 0x00)
                    {
                        break;
                    }

                    countString += 2;
                }

                byte[] byteTemp = new byte[countString - index];
                Array.Copy(rawData, index, byteTemp, 0x00, countString - index);
                string strTemp = Encoding.Unicode.GetString(byteTemp);
                permissionUserInfo.PidTagMemberName = strTemp; // BitConverter.(rawData, nIndex, nCountString - nIndex+1);
                index = countString + 2;

                // PT_LONG  PidTagMemberRights
                byte[] nrights = new byte[4];
                Array.Copy(rawData, index, nrights, 0x00, 0x0004);
                permissionUserInfo.PidTagMemberRights = this.GetRightsByByteArray(ref nrights);
                index += 4;

                // PT_BINARY PidTagEntryId               
                int ncount = 0x02; // COUNT: 16 bits wide
                byte[] countValue = new byte[0x02];
                Array.Copy(rawData, index, countValue, 0x00, ncount);
                int nwide = BitConverter.ToInt16(countValue, 0x0000);
                byte[] nvalue = new byte[nwide];
                index += 2;
                if (nwide > 0)
                {
                    Array.Copy(rawData, index, nvalue, 0x00, nwide - 1);
                }

                index += nwide;
                permissionUserInfo.PidTagEntryId = nvalue;
                permissionUserList.Add(permissionUserInfo);
                --rowCount;
            }

            return true;
        }

        /// <summary>
        /// Get EntryId by user name.
        /// </summary>
        /// <param name="permissionUserName">User name.</param>
        /// <returns>EntryId which is retrieved by server</returns>
        private byte[] GetEntryId(string permissionUserName)
        {
            // Generate the Entry ID for the anonymous and default user.
            if (string.Equals(this.anonymousUser, permissionUserName, StringComparison.OrdinalIgnoreCase) || string.Empty == permissionUserName)
            {
                return new byte[0];
            }
            else
            {
                string distinguishedName = string.Empty;
                if (string.Equals(permissionUserName, this.User2, StringComparison.OrdinalIgnoreCase))
                {
                    distinguishedName = this.AdminUserEssdn + StringNullTerminator;
                }
                else if (string.Equals(permissionUserName, this.User1, StringComparison.OrdinalIgnoreCase))
                {
                    distinguishedName = this.User1Essdn + StringNullTerminator;
                }
                else
                {
                    return null;
                }

                int pidEntryIdLength = 28 + distinguishedName.Length;
                byte[] pidEntryId = new byte[pidEntryIdLength];

                // Create the PidTagEntryId as PermanentEntryID described in section 2.3.8.3, [MS-OXNSPI]
                int i = 0;
                pidEntryId[i] = 0x00;
                i++;
                pidEntryId[i] = 0x00;
                i++;
                pidEntryId[i] = 0x00;
                i++;
                pidEntryId[i] = 0x00;
                i++;

                byte[] providerUID = new byte[16] { 0xDC, 0xA7, 0x40, 0xC8, 0xC0, 0x42, 0x10, 0x1A, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82 };
                Array.Copy(providerUID, 0, pidEntryId, i, 16);
                i += 16;

                byte[] r4 = BitConverter.GetBytes(0x00000001);
                Array.Copy(r4, 0, pidEntryId, i, 4);
                i += 4;

                byte[] displayTypeString = new byte[4] { 0, 0, 0, 0 };
                Array.Copy(displayTypeString, 0, pidEntryId, i, 4);
                i += 4;

                byte[] distinguishedNameBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(distinguishedName);
                Array.Copy(distinguishedNameBytes, 0, pidEntryId, i, distinguishedName.Length);

                return pidEntryId;
            }
        }

        /// <summary>
        /// Set the permissionData array for adding by the permissionUserName and permissionRight
        /// </summary>
        /// <param name="permissionUserName">The user whose permission is for adding</param>
        /// <param name="memberRights">The rights will be assigned to user</param>
        /// <returns>Return the permissionData array for adding user</returns>
        private PermissionData[] SetPermissionDataArrayForAdd(string permissionUserName, uint memberRights)
        {
            PropertyValue[] propertyValues = new PropertyValue[2];
            propertyValues[0] = this.CreateRightsProperty(memberRights);
            propertyValues[1] = this.CreateEntryIdProperty(permissionUserName);

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = 0x01; // AddRow flags 
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;

            return permissionsDataArray;
        }

        /// <summary>
        /// Create the TaggedProperty structure for PidTagMemberRights property.
        /// </summary>
        /// <param name="memberRights">Integer rights</param>
        /// <returns>TaggedProperty indicate permission</returns>
        private TaggedPropertyValue CreateRightsProperty(uint memberRights)
        {
            TaggedPropertyValue rightsProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = 0x6673;
            temp.PropertyType = 0x0003;
            rightsProperty.PropertyTag = temp;
            rightsProperty.Value = BitConverter.GetBytes(memberRights);

            return rightsProperty;
        }

        /// <summary>
        /// Create the TaggedProperty structure for PidTagEntryId property.
        /// </summary>
        /// <param name="permissionUserName">The user name</param>
        /// <returns>TaggedPropertyValue indicate EntryId value</returns>
        private TaggedPropertyValue CreateEntryIdProperty(string permissionUserName)
        {
            TaggedPropertyValue entryIdProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = 0x0fff;
            temp.PropertyType = 0x0102;
            entryIdProperty.PropertyTag = temp;

            entryIdProperty.VarLength = true;
            entryIdProperty.Value = this.GetEntryId(permissionUserName);
            return entryIdProperty;
        }

        /// <summary>
        /// Create the TaggedProperty structure for PidTagMemberId property.
        /// </summary>
        /// <param name="pidTagMemberId">Tag member ID</param>
        /// <returns>TaggedPropertyValue containing PidTagMemberId value</returns>
        private TaggedPropertyValue CreateMemberIdProperty(ulong pidTagMemberId)
        {
            TaggedPropertyValue memberIdProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = 0x6671;
            temp.PropertyType = 0x0014;
            memberIdProperty.PropertyTag = temp;
            memberIdProperty.Value = BitConverter.GetBytes(pidTagMemberId);
            return memberIdProperty;
        }

        /// <summary>
        /// Set the permissionData array for modifying by the PidTagMemberId and PidTagMemberRights.
        /// </summary>
        /// <param name="pidTagMemberId">The pidTagMemberId that specifies the unique identifier that the server generates for each user. </param>
        /// <param name="memberRights">The permission data</param>
        /// <returns>The permissionData array</returns>
        private PermissionData[] SetPermissionDataArrayForModify(ulong pidTagMemberId, uint memberRights)
        {
            PropertyValue[] propertyValues = new PropertyValue[2];

            propertyValues[0] = this.CreateMemberIdProperty(pidTagMemberId);

            propertyValues[1] = this.CreateRightsProperty(memberRights);

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = 0x02; // ModifyRow flags 
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;

            return permissionsDataArray;
        }

        /// <summary>
        /// Set the permissionData array for removing by the PidTagMemberId
        /// </summary>
        /// <param name="pidTagMemberId">The PidTagMemberId</param>
        /// <returns>Return the permissionData Array</returns>
        private PermissionData[] SetPermissionDataArrayForRemove(ulong pidTagMemberId)
        {
            PropertyValue[] propertyValues = new PropertyValue[1];

            propertyValues[0] = this.CreateMemberIdProperty(pidTagMemberId);

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = 0x04; // RemoveRow flags 
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;

            return permissionsDataArray;
        }

        /// <summary>
        /// Get the PidTagMemberId by searching the List
        /// </summary>
        /// <param name="permissionUserList">The total values of the permissions</param>
        /// <param name="permissionUserName">The permissionUserName</param>
        /// <param name="memberIdIsExisted">Indicates whether the PidTagMemberId for the user specified by permissionUserName exists</param>
        /// <returns>Return the matched PidTagMemberId</returns>
        private ulong GetPidTagMemberId(List<PermissionUserInfo> permissionUserList, string permissionUserName, out bool memberIdIsExisted)
        {
            foreach (PermissionUserInfo permissionUserInfo in permissionUserList)
            {
                if (string.Compare(permissionUserName, permissionUserInfo.PidTagMemberName, true) == 0)
                {
                    memberIdIsExisted = true;
                    return permissionUserInfo.PidTagMemberId;
                }
            }

            // The permissionUserName does not exist in the permissions list
            memberIdIsExisted = false;
            return UInt64NULL;
        }

        /// <summary>
        /// Get the user's password.
        /// </summary>
        /// <param name="user">The user whose password to be got.</param>
        /// <returns>The user's password.</returns>
        private string GetUserPassword(string user)
        {
            if (user.Equals(this.User1))
            {
                return this.User1Password;
            }
            else if (user.Equals(this.User2))
            {
                return this.AdminUserPassword;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Create the permissionRight by the PermissionList
        /// </summary>
        /// <param name="permissionList">Permission list</param>
        /// <returns>Permission rights</returns>
        private uint ParsePermissionFromPermissionTypeListToByteFormat(List<PermissionTypeEnum> permissionList)
        {
            uint rights = 0x00000000;
            if (permissionList.Contains(PermissionTypeEnum.ReadAny))
            {
                rights |= 0x00000001;
            }

            if (permissionList.Contains(PermissionTypeEnum.Create))
            {
                rights |= 0x00000002;
            }

            if (permissionList.Contains(PermissionTypeEnum.EditOwned))
            {
                rights |= 0x00000008;
            }

            if (permissionList.Contains(PermissionTypeEnum.DeleteOwned))
            {
                rights |= 0x00000010;
            }

            if (permissionList.Contains(PermissionTypeEnum.EditAny))
            {
                rights |= 0x00000020;
            }

            if (permissionList.Contains(PermissionTypeEnum.DeleteAny))
            {
                rights |= 0x00000040;
            }

            if (permissionList.Contains(PermissionTypeEnum.CreateSubFolder))
            {
                rights |= 0x00000080;
            }

            if (permissionList.Contains(PermissionTypeEnum.FolderOwner))
            {
                rights |= 0x00000100;
            }

            if (permissionList.Contains(PermissionTypeEnum.FolderContact))
            {
                rights |= 0x00000200;
            }

            if (permissionList.Contains(PermissionTypeEnum.FolderVisible))
            {
                rights |= 0x00000400;
            }

            if (permissionList.Contains(PermissionTypeEnum.FreeBusySimple))
            {
                rights |= 0x00000800;
            }

            if (permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed))
            {
                rights |= 0x00001000;
            }

            if (permissionList.Contains(PermissionTypeEnum.Reserved20Permission))
            {
                rights |= 0xFFFFE004;
            }

            return rights;
        }

        /// <summary>
        /// Get right list from rights mask
        /// </summary>
        /// <param name="permissionList">Return value is used to keep the permission list</param>
        /// <param name="rights">Permission mask</param>
        private void ParseRightsFromByteToPermissionList(out List<PermissionTypeEnum> permissionList, uint rights)
        {
            permissionList = new List<PermissionTypeEnum>();
            if ((rights & 0x00000001) == 0x00000001)
            {
                permissionList.Add(PermissionTypeEnum.ReadAny);
            }

            if ((rights & 0x00000002) == 0x00000002)
            {
                permissionList.Add(PermissionTypeEnum.Create);
            }

            if ((rights & 0x00000008) == 0x00000008)
            {
                permissionList.Add(PermissionTypeEnum.EditOwned);
            }

            if ((rights & 0x00000010) == 0x00000010)
            {
                permissionList.Add(PermissionTypeEnum.DeleteOwned);
            }

            if ((rights & 0x00000020) == 0x00000020)
            {
                permissionList.Add(PermissionTypeEnum.EditAny);
            }

            if ((rights & 0x00000040) == 0x00000040)
            {
                permissionList.Add(PermissionTypeEnum.DeleteAny);
            }

            if ((rights & 0x00000080) == 0x00000080)
            {
                permissionList.Add(PermissionTypeEnum.CreateSubFolder);
            }

            if ((rights & 0x00000100) == 0x00000100)
            {
                permissionList.Add(PermissionTypeEnum.FolderOwner);
            }

            if ((rights & 0x00000200) == 0x00000200)
            {
                permissionList.Add(PermissionTypeEnum.FolderContact);
            }

            if ((rights & 0x00000400) == 0x00000400)
            {
                permissionList.Add(PermissionTypeEnum.FolderVisible);
            }

            if ((rights & 0x00001000) == 0x00001000)
            {
                permissionList.Add(PermissionTypeEnum.FreeBusyDetailed);
            }

            if ((rights & 0x00000800) == 0x00000800)
            {
                permissionList.Add(PermissionTypeEnum.FreeBusySimple);
            }

            if ((rights & 0xFFFFE000) == 0xFFFFE000)
            {
                permissionList.Add(PermissionTypeEnum.Reserved20Permission);
            }
        }

        /// <summary>
        /// Get the permissionRight by the Byte Array
        /// </summary>        
        /// <param name="byteRights">Byte Array</param>
        /// <returns>Return the permission Right</returns>
        private uint GetRightsByByteArray(ref byte[] byteRights)
        {
            uint nrights = 0x0000;

            // 19bit reserved.           
            uint rights = (uint)BitConverter.ToInt32(byteRights, 0x0000);

            if ((rights & 0x00000001) == 0x00000001)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000002) == 0x00000002)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000008) == 0x00000008)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000010) == 0x00000010)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000020) == 0x00000020)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000040) == 0x00000040)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000080) == 0x00000080)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000100) == 0x00000100)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000200) == 0x00000200)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000400) == 0x00000400)
            {
                nrights |= rights;
            }

            if ((rights & 0x00001000) == 0x00001000)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000800) == 0x00000800)
            {
                nrights |= rights;
            }

            if ((rights & 0xFFFFE000) == 0xFFFFE000)
            {
                nrights |= rights;
            }

            if ((rights & 0x00000004) == 0x00000004)
            {
                nrights |= rights;
            }

            return nrights;
        }

        /// <summary>
        /// Register the properties to parse
        /// </summary>
        private void RegisterToParses()
        {
            parses = new Dictionary<int, IPropertyInterface>();

            RegisterForParses(0x3001, new DisplayName());
            RegisterForParses(0x6748, new FID());
            RegisterForParses(0x674A, new MessageID());
            RegisterForParses(0x3705, new AttachMethod());
        }

        /// <summary>
        /// Use to  parse and sort the Response buffer 
        /// </summary>
        private struct PermissionUserInfo
        {
            /// <summary>
            /// PidTagEntryId property
            /// </summary>
            public byte[] PidTagEntryId;

            /// <summary>
            /// PidTagMemberId property
            /// </summary>
            public ulong PidTagMemberId;

            /// <summary>
            /// PidTagMemberName property
            /// </summary>
            public string PidTagMemberName;

            /// <summary>
            /// PidTagMemberRights property
            /// </summary>
            public uint PidTagMemberRights;
        }

        #endregion
    }
}