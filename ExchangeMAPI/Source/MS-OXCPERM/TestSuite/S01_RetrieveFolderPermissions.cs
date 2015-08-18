namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the RopGetPermissionsTable operation with parameters.
    /// </summary>
    [TestClass]
    public class S01_RetrieveFolderPermissions : TestSuiteBase
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
        /// This case verifies that the permissions list of a folder can be retrieved.
        /// </summary>
        [TestCategory("MSOXCPERM"), TestMethod()] 
        public void MSOXCPERM_S01_TC01_GetFolderPermissionsList()
        {
            this.OxcpermAdapter.InitializePermissionList();

            // Set the calendar folder type
            FolderTypeEnum folderType = FolderTypeEnum.CalendarFolderType;
            uint responseValue = 0;

            // Set IncludeFreeBusy flag
            RequestBufferFlags bufferflags = new RequestBufferFlags
            {
                IsIncludeFreeBusyFlagSet = true
            };
            List<PermissionTypeEnum> permissionList = new List<PermissionTypeEnum>
            {
                PermissionTypeEnum.FreeBusyDetailed,
                PermissionTypeEnum.FreeBusySimple
            };

            // Add the FreeBusyDetailed and FreeBusySimple to the calendar folder.
            responseValue = OxcpermAdapter.AddPermission(folderType, this.User1, bufferflags, permissionList);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server adds permission successfully.");

            // Get the permissions list on the folder
            permissionList.Clear();
            responseValue = OxcpermAdapter.GetPermission(folderType, this.User1, bufferflags, out permissionList);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R2014");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R2014
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.UINT32SUCCESS,
                responseValue,
                2014,
                @"[In Processing a RopGetPermissionsTable ROP Request] If the user has permission to view the permissions list of the folder, the server returns the permissions list in a RopQueryRows ROP response buffer ([MS-OXCROPS] section 2.2.5.4). ");

            // Verify MS-OXCPERM requirement: MS-OXCPERM_R19
            bool isR19Verified = permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed) && permissionList.Contains(PermissionTypeEnum.FreeBusySimple);
            string returnedPermission = this.ConvertFreeBusyStatusToString(permissionList);

            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCPERM_R19: If the IncludeFreeBusy is set, the server should return {FreeBusyDetailed, FreeBusySimple}, actually " + returnedPermission);
            Site.CaptureRequirementIfIsTrue(
                isR19Verified,
                19,
                @"[In RopGetPermissionsTable ROP Request Buffer] If this flag [IncludeFreeBusy] is set, the server MUST include the values of the FreeBusySimple and FreeBusyDetailed flags of the PidTagMemberRights property in the returned permissions list.");

            // Remove the user from the permissions list that added in this case
            responseValue = OxcpermAdapter.RemovePermission(folderType, this.User1, bufferflags);
            Site.Assert.AreEqual<uint>(TestSuiteBase.UINT32SUCCESS, responseValue, "0 indicates the server removes permission successfully.");
        }

        #endregion Test Cases
    }
}