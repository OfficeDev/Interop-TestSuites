namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Traditional test suite: MS-OXCPERM
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        
        /// <summary>
        /// Indicate the success server response value
        /// </summary>
        protected const uint UINT32SUCCESS = 0x00000000;

        /// <summary>
        /// The anonymous user.
        /// </summary>
        protected const string AnonymousUser = "Anonymous";

        /// <summary>
        /// The default user.
        /// </summary>
        protected const string DefaultUser = "";
        
        /// <summary>
        /// The user configured by "User1Name"
        /// </summary>
        private string user1;

        /// <summary>
        /// The user configured by "AdminUserName"
        /// </summary>
        private string user2;

        /// <summary>
        /// An instance of IMS_OXCPERMAdapter
        /// </summary>
        private IMS_OXCPERMAdapter oxcpermAdapter;

        /// <summary>
        /// Specify whether the case need to do cleanup
        /// </summary>
        private bool needDoCleanup;

        #endregion

        #region Properties

        /// <summary>
        /// Gets the User1
        /// </summary>
        public string User1
        {
            get
            {
                return this.user1;
            }
        }

        /// <summary>
        /// Gets the User2
        /// </summary>
        public string User2
        {
            get
            {
                return this.user2;
            }
        }

        /// <summary>
        /// Gets the instance of IMS_OXCPERMAdapter
        /// </summary>
        public IMS_OXCPERMAdapter OxcpermAdapter
        {
            get
            {
                return this.oxcpermAdapter;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the case need to do cleanup
        /// </summary>
        public bool NeedDoCleanup
        {
            get
            {
                return this.needDoCleanup;
            }

            set
            {
                this.needDoCleanup = value;
            }
        }
        #endregion

        #region Test Suite Initialization

        /// <summary>
        /// Overrides TestClassBase's TestInitialize()
        /// </summary>
        protected override void TestInitialize()
        {
            this.oxcpermAdapter = Site.GetAdapter<IMS_OXCPERMAdapter>();
            this.user1 = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            this.user2 = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            this.needDoCleanup = true;
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            bool transportIsMAPI = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http";
            if (!transportIsMAPI || (transportIsMAPI && Common.IsRequirementEnabled(1184, this.Site)))
            {
                if (this.needDoCleanup)
                {
                    this.OxcpermAdapter.Reset();
                }

                this.OxcpermAdapter.Disconnect();
            }
        }

        #endregion

        #region Protected Method

        /// <summary>
        /// Convert the free/busy status to a string, the format is "{FreeBusyDetailed, FreeBusySimple}" from the permissions list.
        /// </summary>
        /// <param name="permissionList">The permissions list.</param>
        /// <returns>The string contains the free/busy status.</returns>
        protected string ConvertFreeBusyStatusToString(List<PermissionTypeEnum> permissionList)
        {
            string returnedPermission = "{";
            if (permissionList.Contains(PermissionTypeEnum.FreeBusyDetailed))
            {
                returnedPermission = returnedPermission + PermissionTypeEnum.FreeBusyDetailed.ToString();
            }

            if (permissionList.Contains(PermissionTypeEnum.FreeBusySimple))
            {
                returnedPermission = returnedPermission + ", " + PermissionTypeEnum.FreeBusySimple.ToString();
            }

            returnedPermission = returnedPermission + "}";

            return returnedPermission;
        }
        #endregion
    }
}