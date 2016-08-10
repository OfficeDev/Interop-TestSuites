namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-OXCPERM adapter.
    /// </summary>
    public interface IMS_OXCPERMAdapter : IAdapter
    {
        /// <summary>
        /// Read the folder's PidTagSecurityDescriptorAsXml property.
        /// </summary>
        /// <param name="folderType">The folder type specifies the PidTagSecurityDescriptorAsXml property of the folder is read.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        uint ReadSecurityDescriptorProperty(FolderTypeEnum folderType);

        /// <summary>
        /// Get the permission list for a user of the folder
        /// </summary>
        /// <param name="folderType">folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="requestBufferFlags">The request buffer flags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        uint GetPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, out List<PermissionTypeEnum> permissionList);

        /// <summary>
        ///  Add a permission for a user to the permission list of the folder
        /// </summary>
        /// <param name="folderType">folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="requestBufferFlags">The request buffer flags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>      
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>        
        uint AddPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, List<PermissionTypeEnum> permissionList);

        /// <summary>
        /// Modify the permission list for a user of the folder
        /// </summary>
        /// <param name="folderType">folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="requestBufferFlags">The request buffer flags</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>       
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>        
        uint ModifyPermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags, List<PermissionTypeEnum> permissionList);

        /// <summary>
        /// Remove a permission for a user from the permission list of the folder
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="requestBufferFlags">The request buffer flags</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        uint RemovePermission(FolderTypeEnum folderType, string permissionUserName, RequestBufferFlags requestBufferFlags);

        /// <summary>
        /// Check whether the user has the permission to operate the corresponding behavior specified by permission. 
        /// </summary>
        /// <param name="permission">The permission flag specified in PidTagMemberRights</param>
        /// <param name="userName">The user whose permission is specified in permission argument.</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicates error occurs.</returns>
        uint CheckPidTagMemberRightsBehavior(PermissionTypeEnum permission, string userName);

        /// <summary>
        /// The user connects to the server and logons to the mailbox of the user configured by "AdminUserName" in ptfconfig.
        /// </summary>
        /// <param name="userName">The user to logon to the mailbox of the user configured by "AdminUserName" in ptfconfig</param>
        void Logon(string userName);

        /// <summary>
        /// Create a new message in the mail box folder of the user configured by "AdminUserName" by the logon user.
        /// </summary>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        uint CreateMessageByLogonUser();

        /// <summary>
        /// Disconnect the connection with server.
        /// </summary>
        /// <returns>True indicates disconnecting successfully, otherwise false</returns>
        bool Disconnect();

        /// <summary>
        /// Initialize the permission list.
        /// </summary>
        void InitializePermissionList();

        /// <summary>
        /// Check the error code AccessDenied when calling RopQueryRows ROP.
        /// </summary>
        /// <param name="folderType">Folder type</param>
        /// <param name="permissionUserName">The permission user name</param>
        /// <param name="permissionList">The permission list of the folder specified by folderType</param>
        /// <returns>The return value from the server. 0x00000000 indicates success, others indicate error occurs.</returns>
        uint CheckRopQueryRowsErrorCodeAccessDenied(FolderTypeEnum folderType, string permissionUserName, List<PermissionTypeEnum> permissionList);
    }
}