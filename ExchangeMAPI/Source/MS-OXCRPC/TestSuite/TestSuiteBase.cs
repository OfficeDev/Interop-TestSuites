//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A base class contains common methods and fields used by test cases.
    /// </summary>
    [TestClass]
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class TestSuiteBase : TestClassBase
    {
        #region Variable
        /// <summary>
        /// Indicates the value is zero
        /// </summary>
        public const int ZERO = 0;

        /// <summary>
        /// Holds the value that represents the index of first object handle.
        /// </summary>
        protected const int FIRST = 0;

        /// <summary>
        /// Holds the value that represents no session context linking in ulIcxrLink.
        /// </summary>
        protected const uint UlIcxrLinkForNoSessionLink = 0xFFFFFFFF;

        /// <summary>
        /// Holds the value that represents a successful return value for operations
        /// </summary>
        protected const uint ResultSuccess = 0;

        /// <summary>
        /// Indicates value of OpenModeFlags, 2 means create.
        /// </summary>
        protected const uint OpenModeFlags = 2;

        /// <summary>
        /// The instance of the IMS_OXCRPCAdapter
        /// </summary>
        protected IMS_OXCRPCAdapter oxcrpcAdapter;

        /// <summary>
        /// The instance of the IMS_OXCRPCSUTControlAdapter
        /// </summary>
        protected IMS_OXCRPCSUTControlAdapter oxcrpcControlAdapter;

        /// <summary>
        /// User distinguished name (DN)
        /// </summary>
        protected string userDN;

        /// <summary>
        /// Uses this variable to hold the status code of InitializeRPC method
        /// </summary>
        protected bool returnStatus;

        /// <summary>
        /// Uses this variable to hold return value
        /// </summary>
        protected uint returnValue;

        /// <summary>
        /// Uses this variable to hold the return value when CXH is invalid
        /// </summary>
        protected uint returnValueForInvalidCXH;

        /// <summary>
        /// Contains the maximum length of the rgbAuxOut buffer.
        /// </summary>
        protected uint pcbAuxOut;

        /// <summary>
        /// On input, this parameter contains the maximum size of the rgbOut buffer.
        /// On output, this parameter contains the size of ROP response payload, 
        /// including the size of the RPC_HEADER_EXT header in the rgbOut parameter.
        /// </summary>
        protected uint pcbOut;

        /// <summary>
        /// A Session Context Handle to be used with an AsyncEMSMDB interface.
        /// </summary>
        protected IntPtr pacxh;

        /// <summary>
        /// A Session Context Handle to be used with an EMSMDB interface.
        /// </summary>
        protected IntPtr pcxh;

        /// <summary>
        /// An invalid Session Context Handle to be used with an EMSMDB interface.
        /// </summary>
        protected IntPtr pcxhInvalid;

        /// <summary>
        /// A session index value that is associated with the CXH
        /// </summary>
        protected ushort picxr;

        /// <summary>
        /// Contains the time stamp in which the new Session Context was created.
        /// This parameter and ulIcxrLink are used for linking the Session Context created by EcDoConnectEx method with an existing Session Context.
        /// </summary>
        protected uint pulTimeStamp;

        /// <summary>
        /// A valid value for rgwClientVersion, 0x000c, 0x183e, 0x03e8 is specified in [MS-OXCRPC].
        /// </summary>
        protected ushort[] rgwClientVersion = new ushort[3] { 0x000c, 0x183e, 0x03e8 };

        /// <summary>
        /// The minimum client protocol version the server supports 
        /// </summary>
        protected ushort[] rgwBestVersion = new ushort[3];

        /// <summary>
        /// A table contains response handle of ROP commands
        /// </summary>
        protected List<List<uint>> responseSOHTable = new List<List<uint>>();

        /// <summary>
        /// The response of the ROP commands
        /// </summary>
        protected IDeserializable response;

        /// <summary>
        /// A response handle of the ROP commands
        /// </summary>
        protected uint objHandle;

        /// <summary>
        /// Contains the ROP request payload
        /// </summary>
        protected byte[] rgbIn;

        /// <summary>
        /// Contains the ROP response payload
        /// </summary>
        protected byte[] rgbOut;

        /// <summary>
        /// Contains the auxiliary payload buffer
        /// </summary>
        protected byte[] rgbAuxOut;

        /// <summary>
        /// Holds the value represents USE_PER_MDB_REPLID_MAPPING flag that control the behavior of the RopLogon
        /// </summary>
        protected ulong userPrivilege;

        /// <summary>
        /// An unsigned integer represents value of no-used handle or auxInfo for specified RopCommands.
        /// </summary>
        protected uint unusedInfo = 0;

        /// <summary>
        /// An unsigned integer indicates the authentication level for creating RPC binding
        /// </summary>
        protected uint authenticationLevel;

        /// <summary>
        /// Define user name which can be used by client to access SUT.
        /// </summary>
        protected string userName;

        /// <summary>
        /// Define user password which can be used by client to access SUT.
        /// </summary>
        protected string password;

        /// <summary>
        /// An unsigned integer indicates specify authentication services by identifying the security package that provides the service
        /// </summary>
        protected uint authenticationService;

        /// <summary>
        /// A Boolean value indicates whether MAPI HTTP transport is set or not.
        /// </summary>
        protected bool transportIsMAPI;
        #endregion

        #region Test Case Initialization and Cleanup
        /// <summary>
        /// Initializes the test case before running it
        /// </summary>
        protected override void TestInitialize()
        {
            this.Site = TestClassBase.BaseTestSite;
            this.oxcrpcAdapter = this.Site.GetAdapter<IMS_OXCRPCAdapter>();
            this.oxcrpcControlAdapter = this.Site.GetAdapter<IMS_OXCRPCSUTControlAdapter>();

            this.transportIsMAPI = string.Compare(Common.GetConfigurationPropertyValue("TransportSeq", this.Site), "mapi_http", true, System.Globalization.CultureInfo.InvariantCulture) == 0;

            if (!this.transportIsMAPI)
            {
                this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
                this.pcbOut = ConstValues.ValidpcbOut;
                this.rgbIn = new byte[0];
                this.pcxh = IntPtr.Zero;
                this.pcxhInvalid = IntPtr.Zero;
                this.pacxh = IntPtr.Zero;
                this.pulTimeStamp = 0x00000000;
                this.responseSOHTable = new List<List<uint>>();
                this.response = null;
                this.objHandle = 0;
                this.userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site);
                this.userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
                this.password = Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site);

                // Holds the value represents USE_PER_MDB_REPLID_MAPPING flag that control the behavior of the RopLogon
                this.userPrivilege = 0x01000000;

                #region Initializes Server and Client based on specific protocol sequence
                this.authenticationLevel = (uint)Convert.ToInt32(Common.GetConfigurationPropertyValue("RpcAuthenticationLevel", this.Site));
                this.authenticationService = (uint)Convert.ToInt32(Common.GetConfigurationPropertyValue("RPCAuthenticationService", this.Site));
                #endregion
            }
        }

        /// <summary>
        /// Clean up the test case after running it
        /// </summary>
        protected override void TestCleanup()
        {
            if (!this.transportIsMAPI)
            {
                this.HardDeleteMessagesAndSubfolders(FolderIds.Inbox);
                this.HardDeleteMessagesAndSubfolders(FolderIds.SentItems);
            }

            base.TestCleanup();
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Hard delete messages and subfolders under the specified folder.
        /// </summary>
        /// <param name="folderIndex">The id of folder in which subfolders and messages should be deleted</param>
        protected void HardDeleteMessagesAndSubfolders(FolderIds folderIndex)
        {
            #region Client connects with Server
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.returnValue = this.oxcrpcAdapter.EcDoConnectEx(
                ref this.pcxh,
                TestSuiteBase.UlIcxrLinkForNoSessionLink,
                ref this.pulTimeStamp,
                null,
                this.userDN,
                ref this.pcbAuxOut,
                this.rgwClientVersion,
                out this.rgwBestVersion,
                out this.picxr);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoConnectEx should succeed. This call is the precondition for EcDoRpcExt2. '0' is expected to be returned. The returned value is {0}.", this.returnValue);
            #endregion

            #region Logon to mailbox
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopLogon, this.unusedInfo, this.userPrivilege);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoCompression | PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "RopLogon should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue); 
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            uint logonHandle = this.responseSOHTable[TestSuiteBase.FIRST][logonResponse.OutputHandleIndex];
            #endregion

            #region OpenFolder
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopOpenFolder, logonHandle, logonResponse.FolderIds[(int)folderIndex]);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            uint payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            Site.Assert.AreEqual<uint>(0, this.returnValue, "RopOpenFolder should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue); 
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            this.objHandle = this.responseSOHTable[TestSuiteBase.FIRST][openFolderResponse.OutputHandleIndex];
            #endregion

            #region HardDeleteFoldersAndMessages
            this.rgbIn = AdapterHelper.ComposeRgbIn(ROPCommandType.RopHardDeleteMessagesAndSubfolders, this.objHandle, 0);
            this.pcbOut = ConstValues.ValidpcbOut;
            this.pcbAuxOut = ConstValues.ValidpcbAuxOut;
            this.responseSOHTable = new List<List<uint>>();
            payloadCount = 0;

            this.returnValue = this.oxcrpcAdapter.EcDoRpcExt2(
                ref this.pcxh,
                PulFlags.NoXorMagic,
                this.rgbIn,
                ref this.rgbOut,
                ref this.pcbOut,
                null,
                ref this.pcbAuxOut,
                out this.response,
                ref this.responseSOHTable,
                out payloadCount,
                ref this.rgbAuxOut);

            // The returned value 1125 means ecNoDelSubmitMsg, that is, deleting a message that has been submitted for sending is not permitted.
            bool retValue = (this.returnValue == 0) || (this.returnValue == 1125);
            Site.Assert.AreEqual<bool>(true, retValue, "The returned status is {0}. TRUE means that RopHardDeleteMessagesAndSubfolders succeed, and FALSE means that RopHardDeleteMessagesAndSubfolders failed.", retValue);
            #endregion

            #region Client disconnects with Server
            this.returnValue = this.oxcrpcAdapter.EcDoDisconnect(ref this.pcxh);
            Site.Assert.AreEqual<uint>(0, this.returnValue, "EcDoDisconnect should succeed and '0' is expected to be returned. The returned value is {0}.", this.returnValue); 
            #endregion
        }

        /// <summary>
        /// Check the transport. If the transport is MAPI HTTP, issue the inconclusive status.
        /// </summary>
        protected void CheckTransport()
        {
            if (this.transportIsMAPI)
            {
                Site.Assume.Inconclusive("MS-OXCRPC doesn't support MS-OXCMAPIHTTP.");
            }
        }
        #endregion
    }
}