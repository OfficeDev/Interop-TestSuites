namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the RopGetPermissionsTable and RopModifyPermissions operations with parameters.
    /// </summary>
    [TestClass]
    public class S02_ModifyFolderPermissions : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the test suite
        /// </summary>
        /// <param name="testContext">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Reset the test environment
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This case verifies that the permissions list on the folder can be replaced.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S02_TC01_ReplaceFolderPermissions()
        {
            this.OxcpermAdapter.InitializePermissionList();

            uint responseValue = 0;
            uint responseValueExpectedNot0;
            uint responseValueExpected0;
            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>();
            List<PermissionTypeEnum> commonPermission = new List<PermissionTypeEnum>();
            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;
            RequestBufferFlags bufferflags = new RequestBufferFlags
            {
                IsIncludeFreeBusyFlagSet = true
            };

            // Set the ReplaceRows flag, then the AddRow flag is set.
            bufferflags.IsReplaceRowsFlagSet = true;

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.Create);
            responseValueExpected0 = OxcpermAdapter.AddPermission(folderType, this.User1, bufferflags, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValueExpected0, "0 indicates the server adds permission successfully.");
            bufferflags.IsReplaceRowsFlagSet = false;
            responseValueExpected0 = OxcpermAdapter.GetPermission(folderType, this.User1, bufferflags, out commonPermission);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValueExpected0, "0 indicates the server gets permission successfully.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R45");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R45 
            bool isVerifiedR45 = commonPermission.Contains(PermissionTypeEnum.Create);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR45,
                45,
                @"[In RopModifyPermissions ROP Request Buffer] If the ReplaceRows flag is set in the ModifyFlags field, entries can only be added.");

            permissionList.Clear();
            responseValueExpected0 = OxcpermAdapter.GetPermission(folderType, TestSuiteBase.DefaultUser, bufferflags, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValueExpected0, "0 indicates the server gets permission successfully.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1197");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1197
            bool isVerifiedR1197 = commonPermission.Contains(PermissionTypeEnum.Create) && TestSuiteBase.UINT32SUCCESS == responseValueExpected0;
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1197,
                1197,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is set, the server MUST replace all existing entries except the default user entry in the current permissions list with the ones contained in the PermissionsData field.");

            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferflags);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server removes permission successfully.");

            // Get the permission for the User1. Error occurs is expected, for there is no permission for the User1 in the permissions list.
            responseValueExpectedNot0 = OxcpermAdapter.GetPermission(folderType, this.User1, bufferflags, out commonPermission);
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValueExpectedNot0, "Non-zero indicates the user's permission is not got successfully.");

            // Not set the ReplaceRows flag.
            bufferflags.IsReplaceRowsFlagSet = false;
            permissionList.Add(PermissionTypeEnum.Create);
            responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferflags, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server adds permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferflags, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Check the user entry is added to the permissions list
            bool isVerifyR51 = responseValueExpectedNot0 != TestSuiteBase.UINT32SUCCESS && responseValue == TestSuiteBase.UINT32SUCCESS;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R51: PermissionDataFlags is set to AddRow, adding a nonexistent user entry specified by the PidTagEntryId to the permissions list is{0} successful.", isVerifyR51 ? string.Empty : " not");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R51
            Site.CaptureRequirementIfIsTrue(
                isVerifyR51,
                51,
                @"[In PermissionData Structure] [AddRow] The user that is specified by the PidTagEntryId property (section 2.2.4) is added to the permissions list.");

            bool isVerifyR47 = responseValueExpectedNot0 != TestSuiteBase.UINT32SUCCESS && responseValue == TestSuiteBase.UINT32SUCCESS;
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R47: PermissionDataFlags is set to AddRow, adding a nonexistent user entry to the permissions list is{0} successful.", isVerifyR47 ? string.Empty : " not");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R47
            Site.CaptureRequirementIfIsTrue(
                isVerifyR47,
                47,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST add entries in the current permissions list according to the changes specified [PermissionDataFlags is AddRow] in the PermissionsData field [when the PermissionDataFlags flag is set to AddRow].");

            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferflags);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server removes permission successfully.");

            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferflags, out permissionList);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1188
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                1188,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST delete entries in the current permissions list according to the changes specified [PermissionDataFlags is RemoveRow] in the PermissionsData field [when the PermissionDataFlags flag is set to RemoveRow].");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R186
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                186,
                @"[In PermissionData Structure] RemoveRow: The user that is identified by the PidTagMemberId property is deleted from the permissions list.");
        }

        /// <summary>
        /// This case verifies that the permissions control the corresponding behaviors.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S02_TC02_VerifyFolderPermissions()
        {
            this.OxcpermAdapter.InitializePermissionList();

            uint responseValue = 0;

            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>();
            RequestBufferFlags bufferFlag = new RequestBufferFlags();
            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;

            // Here the user configured by "AdminUserName" logons on to the server. 
            // Create 2 messages by the logon user to verify the EditAny or DeleteAny rights in subsequent operations.
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");

            // The ReplaceRows flag is not set.
            bufferFlag.IsReplaceRowsFlagSet = false;

            // Add the user entry in the permissions list to modify the permissions later.
            permissionList.Add(PermissionTypeEnum.Create);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // Logon the user configured by "User1Name" to create a message.
            OxcpermAdapter.Logon(this.User1);

            // The message is created to verify the EditOwned, DeleteOwned, EditAny, DeleteAny rights in subsequent operations.
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");

            #region Check Create permission

            #region Set the Create right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);
            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.Create);
            permissionList.Add(PermissionTypeEnum.FolderVisible);

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R2014");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R2014
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                2014,
                @"[In Processing a RopGetPermissionsTable ROP Request] If the user has permission to view the permissions list of the folder, the server returns the permissions list in a RopQueryRows ROP response buffer ([MS-OXCROPS] section 2.2.5.4). ");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to Create flag, Create flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.Create) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.Create),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.Create, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R140
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                140,
                @"[In PidTagMemberRights Property] If this flag [Create] is set, the server MUST allow the specified user's client to create new Message objects in the folder.");
            #endregion

            #region Disable the Create right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.Create);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to Create flag, Create flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.Create) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.Create),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.Create, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R141
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                141,
                @"[In PidTagMemberRights Property] If this flag [Create] is not set, the server MUST NOT allow the user's client to create new Message objects in the folder.");
            #endregion

            #endregion Check Create permission

            #region Check ReadAny permission

            #region Set the ReadAny right flag and necessary flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.ReadAny);
            permissionList.Add(PermissionTypeEnum.FolderVisible);

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to ReadAny flag, ReadAny flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R52
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R52: The modified permission can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                52,
                @"[In PermissionData Structure] ModifyRow: The existing permissions for the user that is identified by the PidTagMemberId property (section 2.2.5) are modified.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R87
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R87: The modified permission can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            this.Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                87,
                @"[In PidTagMemberRights Property] The PidTagMemberRights property ([MS-OXPROPS] section 2.775) specifies the folder permissions that are granted to the specified user.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R182
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R182: When the ReadAny is disabled, the ReadAny flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                182,
                @"[In Processing a RopModifyPermissions ROP Request] If the user does have permission to modify the folder's properties, the server MUST update the permissions list for the folder according to the PermissionData structures listed in the PermissionsData field of the ROP request buffer, as specified in section 2.2.2.1.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.ReadAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R143
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                143,
                @"[In PidTagMemberRights Property] If this flag [ReadAny] is set, the server MUST allow the specified user's client to read any Message object in the folder.");
            #endregion

            #region Disable the ReadAny right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.ReadAny);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to ReadAny flag, ReadAny flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R182
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R182: When disable the ReadAny flag, the ReadAny flag can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.ReadAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.ReadAny),
                182,
                @"[In Processing a RopModifyPermissions ROP Request] If the user does have permission to modify the folder's properties, the server MUST update the permissions list for the folder according to the PermissionData structures listed in the PermissionsData field of the ROP request buffer, as specified in section 2.2.2.1.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.ReadAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R144
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                144,
                @"[In PidTagMemberRights Property] If this flag [ReadAny] is not set, the server MUST NOT allow the user's client to read Message objects that are owned by other users.");
            #endregion

            #endregion Check ReadAny permission

            #region Check EditOwned permission

            #region Set the EditOwned right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.EditOwned);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.ReadAny);

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to EditOwned, the EditOwned can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.EditOwned) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.EditOwned),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.EditOwned, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R131
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                131,
                @"[In PidTagMemberRights Property] If this flag [EditOwned] is set, the server MUST allow the specified user's client to modify a Message object that was created by that user in the folder.");
            #endregion

            #region Disable the EditOwned right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.EditOwned);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to EditOwned, the EditOwned can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.EditOwned) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.EditOwned),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.EditOwned, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R132
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                132,
                @"[In PidTagMemberRights Property] If this flag [EditOwned] is not set, the server MUST NOT allow the user's client to modify Message objects that were created by that user.");
            #endregion

            #endregion Check EditOwned permission

            #region Check EditAny permission

            #region Disable the EditAny right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Add(PermissionTypeEnum.EditOwned);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.ReadAny);

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to EditAny, the EditAny can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.EditAny) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.EditAny),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.EditAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R125
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                125,
                @"[In PidTagMemberRights Property] If this flag [EditAny] is not set, the server MUST NOT allow the user's client to modify Message objects that are owned by other users.");
            #endregion

            #region Set the EditAny right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();

            // If EditAny flag is set, the EditOwned flag and FolderVisible flag must be set as well.
            permissionList.Add(PermissionTypeEnum.EditAny);
            permissionList.Add(PermissionTypeEnum.EditOwned);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.ReadAny);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            bool isR1187Verified = permissionList.Contains(PermissionTypeEnum.EditAny)
                            && permissionList.Contains(PermissionTypeEnum.EditOwned)
                            && permissionList.Contains(PermissionTypeEnum.FolderVisible);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to EditAny, the EditAny can{0} be retrieved.", isR1187Verified ? string.Empty : " not");

            Site.CaptureRequirementIfIsTrue(
                isR1187Verified,
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.EditAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R124
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                124,
                @"[In PidTagMemberRights Property] If this flag [EditAny] is set, the server MUST allow the specified user's client to modify any Message object in the folder.");
            #endregion

            #endregion Check EditAny permission

            #region Check DeleteOwned permission

            #region Set the DeleteOwned right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();

            // Set FolderVisible to OpenFolder 
            permissionList.Add(PermissionTypeEnum.DeleteOwned);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.ReadAny);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to DeleteOwned, the DeleteOwned can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.DeleteOwned) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.DeleteOwned),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.DeleteOwned, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R127
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                127,
                @"[In PidTagMemberRights Property] If this flag [DeleteOwned] is set, the server MUST allow the specified user's client to delete any Message object that was created by that user in the folder.");
            #endregion

            #region Disable the DeleteOwned right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.DeleteOwned);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.DeleteOwned, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R128
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                128,
                @"[In PidTagMemberRights Property] If this flag [DeleteOwned] is not set, the server MUST NOT allow the user's client to delete Message objects that were created by that user.");
            #endregion

            #endregion Check DeleteOwned permission

            #region Check DeleteAny permission

            #region Set the DeleteAny right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();

            // If DeleteAny flag is set, the DeleteOwned flag and FolderVisible flag must be set as well.
            permissionList.Add(PermissionTypeEnum.DeleteAny);
            permissionList.Add(PermissionTypeEnum.DeleteOwned);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.ReadAny);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            isR1187Verified = permissionList.Contains(PermissionTypeEnum.DeleteAny) && permissionList.Contains(PermissionTypeEnum.DeleteOwned);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to DeleteAny, the DeleteAny can{0} be retrieved.", isR1187Verified ? string.Empty : " not");

            Site.CaptureRequirementIfIsTrue(
                isR1187Verified,
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.DeleteAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R121
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                121,
                @"[In PidTagMemberRights Property] If this flag [DeleteAny] is set, the server MUST allow the specified user's client to delete any Message object in the folder.");
            #endregion

            #region Disable the DeleteAny right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.DeleteAny);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.DeleteAny, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R122
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                122,
                @"[In PidTagMemberRights Property] If this flag [DeleteAny] is not set, the server MUST NOT allow the user's client to delete Message objects that are owned by other users.");
            #endregion

            #endregion Check DeleteAny permission

            #region Check CreateSubFolder permission

            #region Set the CreateSubFolder right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            permissionList.Add(PermissionTypeEnum.CreateSubFolder);

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to CreateSubFolder, the CreateSubFolder can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.CreateSubFolder) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.FolderVisible) && permissionList.Contains(PermissionTypeEnum.CreateSubFolder),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.CreateSubFolder, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R118
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                118,
                @"[In PidTagMemberRights Property] If this flag [CreateSubFolder] is set, the server MUST allow the specified user's client to create new folders within the folder.");
            #endregion

            #region Disable the CreateSubFolder right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);
            permissionList.Remove(PermissionTypeEnum.CreateSubFolder);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.CreateSubFolder, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R119
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                119,
                @"[In PidTagMemberRights Property] If this flag [CreateSubFolder] is not set, the server MUST NOT allow the user's client to create new folders within the folder.");
            #endregion

            #endregion Check CreateSubFolder permission

            #region Check FolderOwner permission

            #region Set the FolderOwner right flag and necessary right flags, keep other right flags unchanged.
           
                OxcpermAdapter.Logon(this.User2);

                permissionList.Clear();
                permissionList.Add(PermissionTypeEnum.FolderOwner);
                permissionList.Add(PermissionTypeEnum.FolderVisible);

                // If FolderOwner flag is set, FolderVisible flag must be set as well.
                permissionList.Add(PermissionTypeEnum.FolderVisible);
                responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
                Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

                permissionList.Clear();
                responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
                Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

                // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
                isR1187Verified = permissionList.Contains(PermissionTypeEnum.FolderOwner) && permissionList.Contains(PermissionTypeEnum.FolderVisible);
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to FolderOwner, the FolderOwner can{0} be retrieved.", isR1187Verified ? string.Empty : " not");

                Site.CaptureRequirementIfIsTrue(
                    isR1187Verified,
                    1187,
                    @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");
                if (Common.IsRequirementEnabled(115002, this.Site))
                {
                OxcpermAdapter.Logon(this.User1);

                // Check the user has the set permission to handle the corresponding behavior.
                responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FolderOwner, this.User1);

                // Verify MS-OXCPERM requirement: MS-OXCPERM_R115002
                Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.UINT32SUCCESS,
                    responseValue,
                    115002,
                    "[In Appendix A: Product Behavior] When flag FolderOwner is set, the Implementation does allow the specified user's client to modify properties, including the folder permissions, that are set on the folder itself. (Exchange 2010 and above follow this behavior.)");
                }
                
            #endregion

            #region Disable the FolderOwner right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.FolderOwner);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FolderOwner, this.User1);

            if (Common.IsRequirementEnabled(116001, this.Site))
            {
                // Verify MS-OXCPERM requirement: MS-OXCPERM_R116001
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    TestSuiteBase.UINT32SUCCESS,
                    responseValue,
                    116001,
                    @"[In Appendix A: Product Behavior] When this flag FolderOwner is not set, the Implementation does not allow the specified user's client to modify the folder's properties. (Exchange 2013 and above follow this behavior.)");

            }

            if (Common.IsRequirementEnabled(116002, this.Site))
            {
                // Verify MS-OXCPERM requirement: MS-OXCPERM_R116002
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    TestSuiteBase.UINT32SUCCESS,
                    responseValue,
                    116002,
                    @"[In Appendix A: Product Behavior] When this flag FolderOwner is not set, the Implementation does allow the specified user's client to modify the folder's properties. <4> Section 2.2.7:  Exchange 2007 and Exchange 2010 allow the properties of a folder to be modified when the FolderOwner flag is not set.");

            }
                // Verify MS-OXCPERM requirement: MS-OXCPERM_R181
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                181,
                "[In Processing a RopModifyPermissions ROP Request] If the user does not have permission to modify the folder's properties, the server MUST return the AccessDenied (0x80070005) error code in the ReturnValue field of the ROP response buffer.");
            #endregion

            #endregion Check FolderOwner permission

            #region Check FolderVisible permission

            #region Set the FolderVisible right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to FolderVisible, the FolderVisible can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.FolderVisible) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.FolderVisible),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FolderVisible, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R103
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                103,
                @"[In PidTagMemberRights Property] If this flag [FolderVisible] is set, the server MUST allow the specified user's client to retrieve the folder's permissions list, as specified in section 3.1.4.1, to see the folder in the folder hierarchy table.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1120
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                1120,
                @"[In PidTagMemberRights Property] If this flag [FolderVisible] is set, the server MUST allow the specified user's client to retrieve the folder's permissions list, as specified in section 3.1.4.1, to open the folder.");

            #endregion

            #region Disable the FolderVisible right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FolderVisible, this.User1);
            bool containsFolderVisible = permissionList.Contains(PermissionTypeEnum.FolderVisible);
            Site.Assert.AreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                "If the FolderVisible flag is not set, the server MUST deny the specified user's client to open the folder. FolderVisible is {0} set.",
                containsFolderVisible ? string.Empty : "not");

            #endregion

            #endregion Check FolderVisible permission

            #region Check FreeBusySimple permission

            #region Set the FreeBusySimple right flag, keep other right flag unchanged.

            OxcpermAdapter.Logon(this.User2);
            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.FreeBusySimple);
            bufferFlag.IsIncludeFreeBusyFlagSet = true; // Set IncludeFreeBusy flag
            folderType = FolderTypeEnum.CalendarFolderType;

            // Add the user entry in the permissions list
            responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to FreeBusySimple, the FreeBusySimple can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.FreeBusySimple) ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                permissionList.Contains(PermissionTypeEnum.FreeBusySimple),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FreeBusySimple, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R97
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                97,
                @"[In PidTagMemberRights Property] If this flag [FreeBusySimple] is set, the server MUST allow the specified user's client to retrieve brief information about the appointments on the calendar through the Availability Web Service Protocol, as specified in [MS-OXWAVLS].");
            #endregion

            #region Disable the FreeBusySimple right flag, keep other right flag unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to FreeBusySimple, the FreeBusySimple can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.FreeBusySimple) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.FreeBusySimple),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FreeBusySimple, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R98
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                98,
                @"[In PidTagMemberRights Property] If this flag [FreeBusySimple] is not set, the server MUST NOT allow the specified user's client to retrieve information through the Availability Web Service Protocol.<5>");
            #endregion

            #endregion Check FreeBusySimple permission

            #region Check FreeBusyDetailed permission

            #region Set the FreeBusyDetailed right flag and necessary right flags, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.FreeBusyDetailed);

            // If FreeBusyDetailed flag is set, FreeBusySimple flag must be set as well. 
            permissionList.Add(PermissionTypeEnum.FreeBusySimple);
            bufferFlag.IsIncludeFreeBusyFlagSet = true;
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            isR1187Verified = permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed) && permissionList.Contains(PermissionTypeEnum.FreeBusySimple);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is modified to FreeBusyDetailed, the FreeBusyDetailed can{0} be retrieved.", isR1187Verified ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                isR1187Verified,
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R52
            bool isR52Verified = permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed) && permissionList.Contains(PermissionTypeEnum.FreeBusySimple);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R52: The modified permission is{0} retrieved.", isR52Verified ? string.Empty : " not");
            Site.CaptureRequirementIfIsTrue(
                isR52Verified,
                52,
                @"[In PermissionData Structure] ModifyRow: The existing permissions for the user that is identified by the PidTagMemberId property (section 2.2.5) are modified.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R40
            bool isR40Verified = permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed) && permissionList.Contains(PermissionTypeEnum.FreeBusySimple);
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R40: The IncludeFreeBusy flag does{0} acknowledge the FreeBusySimple and FreeBusyDetailed flags.", isR40Verified ? string.Empty : " not");
            this.Site.CaptureRequirementIfIsTrue(
                isR40Verified,
                40,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [IncludeFreeBusy] is set, the server MUST apply the settings of the FreeBusySimple and FreeBusyDetailed flags of the PidTagMemberRights property when modifying the permissions of the Calendar folder.");

            OxcpermAdapter.Logon(this.User1);

            // Check the user has the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FreeBusyDetailed, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R91
            Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                91,
                @"[In PidTagMemberRights Property] If this flag [FreeBusyDetailed] is set, the server MUST allow the specified user's client to retrieve detailed information about the appointments on the calendar through the Availability Web Service Protocol, as specified in [MS-OXWAVLS].");
            #endregion

            #region Disable the FreeBusyDetailed right flag, keep other right flags unchanged.

            OxcpermAdapter.Logon(this.User2);

            permissionList.Remove(PermissionTypeEnum.FreeBusyDetailed);
            responseValue = OxcpermAdapter.ModifyPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server modifies permission successfully.");

            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1187
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1187: The permission is not modified to FreeBusyDetailed, the FreeBusyDetailed can{0} be retrieved.", permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed) ? string.Empty : " not");
            Site.CaptureRequirementIfIsFalse(
                permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed),
                1187,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST update entries in the current permissions list according to the changes specified [PermissionDataFlags is ModifyRow] in the PermissionsData field [when the PermissionDataFlags flag is set to ModifyRow].");

            OxcpermAdapter.Logon(this.User1);

            // Check the user doesn't have the set permission to handle the corresponding behavior.
            responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(PermissionTypeEnum.FreeBusyDetailed, this.User1);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R92
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                92,
                @"[In PidTagMemberRights Property] If this flag [FreeBusyDetailed] is not set, the server MUST NOT allow the specified user's client to see these details.");

            #endregion

            #endregion Check FreeBusyDetailed permission

            // That the above positive requirements can be passed indicates that the client can allow the access based on the permissions that the right flags are set.
            Site.CaptureRequirement(
                173,
                @"[In Accessing a Folder] When a client sends a request to the server to access a folder, as specified in [MS-OXCFOLD], the server MUST allow the request based on the permissions list for the folder and the user credentials that the client provided enable access when making the request.");

            // That the above negative requirements can be passed indicates that the client can deny the access based on the permissions that the right flags are not set.
            Site.CaptureRequirement(
                1192,
                @"[In Accessing a Folder] When a client sends a request to the server to access a folder, as specified in [MS-OXCFOLD], the server MUST deny the request if the permissions list for the folder and the user credentials that the client provided disable access when making the request.");

            // That the above requirements can be passed indicates that the user's permissions are consistent with the permissions in the permissions list for specified folder.
            this.Site.CaptureRequirement(
                175,
                @"[In Accessing a Folder] Specific user permissions: If the user is included in the permissions list, either explicitly or through membership in a group that is included in the permissions list, the server MUST apply the permissions that have been set for that user. ");

            // That the above requirements can be passed indicates that the server determines whether the user is included in the permissions list and applies the folder permissions for the user.
            this.Site.CaptureRequirement(
                2016,
                @"[In Accessing a Folder] The server determines whether the user identified by the user credentials is included in the permissions list and then applies the folder permissions [Specific user permissions, Default user permissions, Anonymous user permissions] for that user.");

            OxcpermAdapter.Logon(this.User2);

            // Remove the user permissions in the permission list on the common folder.
            bufferFlag.IsReplaceRowsFlagSet = false;
            bufferFlag.IsIncludeFreeBusyFlagSet = true;
            folderType = FolderTypeEnum.CommonFolderType;
            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferFlag);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server removes permission successfully.");

            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1188
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                1188,
                @"[In RopModifyPermissions ROP Request Buffer] If this flag [ReplaceRows] is not set, the server MUST delete entries in the current permissions list according to the changes specified [PermissionDataFlags is RemoveRow] in the PermissionsData field [when the PermissionDataFlags flag is set to RemoveRow].");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R186
            Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                186,
                @"[In PermissionData Structure] RemoveRow: The user that is identified by the PidTagMemberId property is deleted from the permissions list.");

            // Remove the user permission in the permissions list on the Calendar folder.
            folderType = FolderTypeEnum.CalendarFolderType;
            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferFlag);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");
        }

        /// <summary>
        /// This case verifies that the permissions for the default user will be applied for a user if his credentials are not included in the permissions list.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S02_TC03_CheckDefaultPermissionUsedForCredentialUser()
        {
            this.OxcpermAdapter.InitializePermissionList();

            Site.Assume.IsTrue(Common.IsRequirementEnabled(115002, this.Site) || Common.IsRequirementEnabled(116001, this.Site), "This case runs only when the implementation supports the FolderOwner permission.");
            uint responseValue = 0;
            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>();
            RequestBufferFlags bufferFlag = new RequestBufferFlags
            {
                IsIncludeFreeBusyFlagSet = true
            };
            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;

            // Here the user configured by "AdminUserName" logons on to the server. 
            // Create 2 messages by the logon user to verify the EditAny or DeleteAny rights in subsequent operations.
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");

            // The ReplaceRows flag is not set.
            bufferFlag.IsReplaceRowsFlagSet = false;

            // Add the user entry in the permissions list to modify the permissions later.
            permissionList.Clear();
            permissionList.Add(PermissionTypeEnum.Create);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // Logon the user configured by "User1Name" to create a message.
            OxcpermAdapter.Logon(this.User1);

            // The message is created to verify the EditOwned, DeleteOwned, EditAny, DeleteAny rights in subsequent operations.
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");
            responseValue = OxcpermAdapter.CreateMessageByLogonUser();
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server create a message successfully.");

            OxcpermAdapter.Logon(this.User2);
            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferFlag);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // The user permission entry is not in the permissions list, the server considers it as default user. 
            // So the user has no his own permissions.
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            // Modify the default user's permission
            permissionList.Clear();
            permissionList = new List<PermissionTypeEnum>
                {
                    PermissionTypeEnum.ReadAny,
                    PermissionTypeEnum.Create,
                    PermissionTypeEnum.EditOwned,
                    PermissionTypeEnum.DeleteOwned,
                    PermissionTypeEnum.EditAny,
                    PermissionTypeEnum.DeleteAny,
                    PermissionTypeEnum.CreateSubFolder,
                    PermissionTypeEnum.FolderVisible
                };
            responseValue = OxcpermAdapter.ModifyPermission(folderType, TestSuiteBase.DefaultUser, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            // Get the default user's permission
            List<PermissionTypeEnum> defaultUserPermission = new List<PermissionTypeEnum>();
            responseValue = OxcpermAdapter.GetPermission(folderType, TestSuiteBase.DefaultUser, bufferFlag, out defaultUserPermission);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server gets permission successfully.");

            OxcpermAdapter.Logon(this.User1);

            // Add all the right flags in the list.
            List<PermissionTypeEnum> allRightFlags = new List<PermissionTypeEnum>
                {
                    PermissionTypeEnum.ReadAny,
                    PermissionTypeEnum.Create,
                    PermissionTypeEnum.EditOwned,
                    PermissionTypeEnum.DeleteOwned,
                    PermissionTypeEnum.EditAny,
                    PermissionTypeEnum.DeleteAny,
                    PermissionTypeEnum.CreateSubFolder,
                    PermissionTypeEnum.FolderOwner,
                    PermissionTypeEnum.FolderVisible
                };

            // Check the provided user credentials for a user that is not in the permissions list has the permissions of the default user.
            foreach (PermissionTypeEnum permission in allRightFlags)
            {
                responseValue = OxcpermAdapter.CheckPidTagMemberRightsBehavior(permission, this.User1);
                bool permissionFlagIsSet = defaultUserPermission.Contains(permission);
                bool permissionIsAllowed = responseValue == TestSuiteBase.UINT32SUCCESS;
                if ((permissionFlagIsSet && permissionIsAllowed) || (!permissionFlagIsSet && !permissionIsAllowed))
                {
                    Site.Log.Add(
                        LogEntryKind.Comment,
                        "Verify the default user permission. The permission flag: {0} is {1} set, the corresponding behavior is {2} allowed.",
                        permission.ToString(),
                        permissionFlagIsSet ? string.Empty : "not",
                        permissionIsAllowed ? string.Empty : "not");
                }
                else
                {
                    Site.Assert.Fail(
                        "Verify the default user permission. The permission flag: {0} is {1} set, the corresponding behavior is {2} allowed.",
                        permission.ToString(),
                        permissionFlagIsSet ? string.Empty : "not",
                        permissionIsAllowed ? string.Empty : "not");
                }
            }

            // Above steps can verify the behaviors. If each the permission and its corresponding behavior are consistent, it can run here, else fail for the inconsistent permission and behavior.
            this.Site.CaptureRequirement(
                2016,
                @"[In Accessing a Folder] The server determines whether the user identified by the user credentials is included in the permissions list and then applies the folder permissions [Specific user permissions, Default user permissions, Anonymous user permissions] for that user.");

            // Above steps can verify the behaviors. If each the permission and its corresponding behavior are consistent, it can run here, else fail for the inconsistent permission and behavior.
            this.Site.CaptureRequirement(
                176,
                @"[In Accessing a Folder] Default user permissions: If the user is not included in the permissions list, the server MUST apply the permissions that have been set in the default user entry of the permissions list.");
        }
        #endregion Test Cases
    }
}