namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the error code or negative behaviors.
    /// </summary>
    [TestClass]
    public class S03_NegativeOrErrorValidation : TestSuiteBase
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
        /// This case verifies that the permission flags are not allowed in the Open Specification.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S03_TC01_SetUnexpectedPermissionFlags()
        {
            this.OxcpermAdapter.InitializePermissionList();

            uint getPermissionResponseValue = 0;
            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>();
            RequestBufferFlags bufferFlag = new RequestBufferFlags();
            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;

            // Set the right flags that are not specified in the Open Specification.
            permissionList.Add(PermissionTypeEnum.Reserved20Permission);
            uint responseValueSetPermission = OxcpermAdapter.AddPermission(folderType, this.User1, bufferFlag, permissionList);

            permissionList.Clear();
            getPermissionResponseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferFlag, out permissionList);

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R1093: Verify the unexpected permission flags specified in [MS-OXCPERM] section 2.2.7 can't be set. The test case expects that if set the unexpected permission flags, return value is not success or return value is success but actually the flags are not set.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R1093
            // If the server fails to set the unexpected flags or the permissions set for the user is empty, it indicates that the server doesn't set the permission flags that are not specified in the Open Specification.
            bool isR1093Verified = responseValueSetPermission != TestSuiteBase.UINT32SUCCESS || (getPermissionResponseValue == TestSuiteBase.UINT32SUCCESS && permissionList.Count == 0);
            Site.CaptureRequirementIfIsTrue(
                isR1093Verified,
                1093,
                @"[In PidTagMemberRights Property] The client and server MUST NOT set any other flags [except ReadAny, Create, EditOwned, DeleteOwned, EditAny, DeleteAny, CreateSubFolder, FolderOwner, FolderContact, FolderVisible, FreeBusySimple, FreeBusyDetailed].");
        }

        /// <summary>
        /// This case verifies the ecNotImplemented error code (0x80040102) when calling RopOpenStream.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S03_TC02_RetrieveFolderPermissionNotImplementedError()
        {
            this.OxcpermAdapter.InitializePermissionList();

            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;
            const uint NotImplemented = 0x80040102;

            uint responseValue = OxcpermAdapter.ReadSecurityDescriptorProperty(folderType);

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R183
            Site.CaptureRequirementIfAreEqual<uint>(
                NotImplemented,
                responseValue,
                183,
                @"[In Processing a Request for PidTagSecurityDescriptorAsXml Property] When the server receives a RopOpenStream ROP request ([MS-OXCROPS] section 2.2.9.1) on the PidTagSecurityDescriptorAsXml property ([MS-XWDVSEC] section 2.2.2) of the folder, the server MUST return an error code of ecNotImplemented rather than satisfying the RopOpenStream ROP request.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R152
            Site.CaptureRequirementIfAreEqual<uint>(
                NotImplemented,
                responseValue,
                152,
                @"[In Retrieving Folder Permissions] The server MUST return an error code of ecNotImplemented instead of satisfying the RopOpenStream ROP request.");

            this.NeedDoCleanup = false;
        }

        /// <summary>
        /// This case verifies the AccessDenied (0x80070005) when calling RopQueryRows to retrieve the permission list if the permission is insufficient.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()]
        public void MSOXCPERM_S03_TC03_VerifyRopQueryRowsErrorCodeAccessDenied()
        {
            this.OxcpermAdapter.InitializePermissionList();

            Site.Assume.IsTrue(Common.IsRequirementEnabled(115002, this.Site)||Common.IsRequirementEnabled(116001, this.Site), "This case runs only when the implementation supports the FolderOwner permission.");
            const uint AccessDenied = 0x80070005;
            FolderTypeEnum folderType = FolderTypeEnum.CommonFolderType;
            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>();
            RequestBufferFlags bufferFlag = new RequestBufferFlags();

            // Add the user entry in the permissions list to modify the permissions later.
            permissionList.Add(PermissionTypeEnum.FolderOwner);
            permissionList.Add(PermissionTypeEnum.FolderVisible);
            uint responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferFlag, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server response successfully.");

            OxcpermAdapter.Logon(this.User1);
            permissionList.Clear();

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R2007: Verify the AccessDenied (0x80070005) when calling RopQueryRows to retrieve the permission list if the permission is insufficient.");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R2007
            responseValue = OxcpermAdapter.CheckRopQueryRowsErrorCodeAccessDenied(folderType, this.User1, permissionList);
            Site.CaptureRequirementIfAreEqual<uint>(
                AccessDenied,
                responseValue,
                2007,
                "[In Processing a RopGetPermissionsTable ROP Request] If the user does not have permission to view the permissions list of the folder, the server returns 0x80070005 (AccessDenied) in the ReturnValue field of the RopQueryRows ROP response buffer.");
        }
        #endregion Test Cases
    }
}