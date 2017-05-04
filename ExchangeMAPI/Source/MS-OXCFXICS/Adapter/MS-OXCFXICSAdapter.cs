namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXCFXICS
    /// </summary>
    public partial class MS_OXCFXICSAdapter : ManagedAdapterBase, IMS_OXCFXICSAdapter
    {
        #region Content definition

        /// <summary>
        /// The GID length.
        /// </summary>
        private const int GidLength = 22;

        /// <summary>
        /// Domain configure property.
        /// </summary>
        private const string Domain = "domain";

        /// <summary>
        /// The OxcropsClient instances.
        /// </summary>
        private Dictionary<int, OxcropsClient> oxcropsClient = new Dictionary<int, OxcropsClient>();

        /// <summary>
        /// Store ICS state value and index.
        /// </summary>
        private Dictionary<int, ICSStateData> icsStateContainer;

        /// <summary>
        /// Record the current permission for user.
        /// </summary>
        private PermissionLevels currentPermission;

        /// <summary>
        /// Record the current copyflag.
        /// </summary>
        private CopyToCopyFlags currentCopyFlag;

        /// <summary>
        /// Inbox folder id.
        /// </summary>
        private ulong inboxFolderId;

        /// <summary>
        /// Transfer Buffer.
        /// </summary>
        private List<byte[]> totalTransferBufferList = new List<byte[]>();

        /// <summary>
        /// Store the map of property name and property tag.
        /// </summary>
        private Dictionary<string, PropertyTag> propertyTagsDictionary = new Dictionary<string, PropertyTag>();

        /// <summary>
        /// Store the map of tagged property value and name.
        /// </summary>
        private Dictionary<string, TaggedPropertyValue> taggedPropertyValuesDictionary = new Dictionary<string, TaggedPropertyValue>();

        /// <summary>
        /// Store values got from get properties specific.
        /// </summary>
        private Dictionary<string, byte[]> propertyValuesSpecific = new Dictionary<string, byte[]>();

        /// <summary>
        /// Logon handle.
        /// </summary>
        private uint logonHandleOut;

        /// <summary>
        /// Local replica GUID.
        /// </summary>
        private Guid localReplicaGuid = Guid.NewGuid();

        /// <summary>
        /// Server replica GUID.
        /// </summary>
        private Guid serverReplicaGuid;

        /// <summary>
        /// The first localId value allotted by server.
        /// </summary>
        private byte[] localId = new byte[6];

        /// <summary>
        /// Identify how many local ids have been used.
        /// </summary>
        private int localIdOffSet;

        /// <summary>
        /// Container of object handle.
        /// </summary>
        private Dictionary<int, uint> handleContainer;

        /// <summary>
        /// Container of object Id.
        /// </summary>
        private Dictionary<int, ulong> objectIdContainer;

        /// <summary>
        /// Container of Transfer Data.
        /// </summary>
        private Dictionary<int, byte[]> streamBufferContainer;

        /// <summary>
        /// The list of out handles in the ROP response.
        /// </summary>
        private List<uint> responseSOHs;

        /// <summary>
        /// The type of the current FastTransferStream
        /// </summary>
        private FastTransferStreamType streamType;

        /// <summary>
        /// Their folders that the client has no permission to access
        /// </summary>
        private bool existNoPermissionFolder;

        /// <summary>
        /// Stores CNSet read.
        /// </summary>
        private Dictionary<int, REPLGUID_IDSET> cnsetRead;

        /// <summary>
        /// Stores CNSet seen.
        /// </summary>
        private Dictionary<int, REPLGUID_IDSET> cnsetSeen;

        /// <summary>
        /// Stores CNSet seen FAI.
        /// </summary>
        private Dictionary<int, REPLGUID_IDSET> cnsetSeenFAI;

        /// <summary>
        /// Public folder id.
        /// </summary>
        private ulong publicFolderId;

        /// <summary>
        /// Id of the folder which is parent of public folder.
        /// </summary>
        private ulong rootFolderId;

        /// <summary>
        /// Indicate whether needs to do the cleanup work
        /// </summary>
        private bool needDoCleanup = true;

        /// <summary>
        /// Indicate whether the second SUT server is connected or not.
        /// </summary>
        private bool isSecondSUTServerConnected;

        /// <summary>
        /// Record the last conflict information.
        /// </summary>
        private ConflictInfo lastConflictInfo = new ConflictInfo();

        /// <summary>
        /// Stores the message id in FastTransferSourceCopyMessages request buffer
        /// </summary>
        private long messageIdForFastTransferSourceCopyMessages;

        /// <summary>
        /// Stores the copy flag in FastTransferSourceCopyMessages request buffer
        /// </summary>
        private RopFastTransferSourceCopyMessagesCopyFlags copyFlagForFastTransferSourceCopyMessages = RopFastTransferSourceCopyMessagesCopyFlags.Unused1;

        /// <summary>
        /// Stores the ROPFastTransferOperation for further using
        /// </summary>
        private EnumFastTransferOperation previousOperation;

        /// <summary>
        /// Stores the propertyTags in FastTransferSourceCopyProperties or FastTransferSourceCopyTo request buffer
        /// </summary>
        private PropertyTag[] propertyTags;

        /// <summary>
        /// Stores the level in FastTransferSourceCopyProperties
        /// </summary>
        private bool copySubObjects;

        /// <summary>
        /// Record the current number of rows in the soft delete contents table.
        /// </summary>
        private int currentSoftDeleteRowCount;

        /// <summary>
        /// Stores the SynchronizationFlag in SynchronizationConfigure
        /// </summary>
        private SynchronizationFlag synchroniztionFlag = SynchronizationFlag.None;

        /// <summary>
        /// Stores the propertyTag in SynchronizationConfigure request buffer
        /// </summary>
        private PropertyTag propertyTagForConfigure;

        /// <summary>
        /// Record the import flag.
        /// </summary>
        private ImportFlag importFlag = ImportFlag.InvalidParameter;

        /// <summary>
        /// Record whether the OrderByDeliveryTime SynchronizationExtraFlag flag in RopSynchronizationConfigure ROP is not set.
        /// </summary>
        private bool isOrderByDeliveryTimeExtraFlagNotSet;

        /// <summary>
        /// Record the previous result of FastTransferSourceGetBuffer.
        /// </summary>
        private RopResult previousGetBufferResult;

        /// <summary>
        /// Record the import PidTagChangeKey value.
        /// </summary>
        private byte[] importPidTagChangeKeyValue = new byte[22];

        /// <summary>
        /// Record whether the last changed is made by server.
        /// </summary>
        private bool lastChangeMadeByServer;

        /// <summary>
        /// Record whether the last changed is made by client.
        /// </summary>
        private bool lastChangeMadeByClient;

        /// <summary>
        /// Record whether there is not importMessageChange ROP in test case.
        /// </summary>
        private bool isNonImportMessageChangeOperation;

        /// <summary>
        /// Record the server called importMessageChange ROP times.
        /// </summary>
        private int importMessageTimes;

        /// <summary>
        /// Record the previous step verifid result when the test case have more than one steps and need know the last step verifed result.
        /// </summary>
        private bool previousStepVerifiedResult;

        /// <summary>
        /// Record the RPC return value of the ROP call.
        /// </summary>
        private uint ropResult;
        #endregion

        #region Test MetaTagIdsetDeleted property related.

        /// <summary>
        /// Record whether has executed the SynchronizationImportDeletes method.
        /// </summary>
        private bool hasExecuteSynchronizationImportDeletes;

        /// <summary>
        /// Record the folder Ids which has deleted in the SynchronizationImportDeletes method.
        /// </summary>
        private ulong delObjId;

        /// <summary>
        /// Record the delete flags in the SynchronizationImportDeletes method.
        /// </summary>
        private byte importDelFlag = 0x10;

        /// <summary>
        /// Record message change information.
        /// </summary>
        private MessageChangeFull messageChangeFull;

        /// <summary>
        /// Record the folderCopyFlag in FastTransferSourceCopyFolder request buffer.
        /// </summary>
        private CopyFolderCopyFlags folderCopyFlag;

        /// <summary>
        /// Record whether the prior ROP operation is ROPSynchronizationImportMessageMove or not.
        /// </summary>
        private bool isImportMessageMoveROP;

        #endregion

        /// <summary>
        /// This method converts a string to bytes array
        /// </summary>
        /// <param name="str">A string to convert</param>
        /// <returns>Return the converted byte array</returns>
        public static byte[] ConvertStringToBytes(string str)
        {
            char[] charSeparators = { ',' };
            string[] strArray = str.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
            byte[] bytes = new byte[strArray.Length];
            for (int i = 0; i < strArray.Length; i++)
            {
                bytes[i] = byte.Parse(strArray[i].Trim().Substring(2), System.Globalization.NumberStyles.HexNumber);
            }

            return bytes;
        }

        #region Rop Interfaces implement

        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXCFXICS";
            Common.MergeConfiguration(testSite);
            this.InitialVariables();
            this.oxcropsClient.Add(1, new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site)));
            if (Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site) != string.Empty)
            {
                this.oxcropsClient.Add(2, new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site)));
            }

            AdapterHelper.Site = testSite;
        }

        /// <summary>
        /// This method is used to check the second system under test whether is online or not.
        /// </summary>
        /// <param name="isSecondSUTOnline">Check second SUT is online or not</param>
        public void CheckSecondSUTOnline(out bool isSecondSUTOnline)
        {
            isSecondSUTOnline = !Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site).Equals(string.Empty);

            if (!isSecondSUTOnline)
            {
                this.needDoCleanup = false;
                Site.Assert.Inconclusive("This case runs only when the second system under test exists.");
            }
        }

        /// <summary>
        /// Reset the adapter.
        /// </summary>
        public override void Reset()
        {
            try
            {
                if (this.needDoCleanup)
                {
                    this.RevertPublicFolder(LogonFlags.PublicFolder);
                    foreach (KeyValuePair<int, OxcropsClient> oxcrop in this.oxcropsClient)
                    {
                        if (oxcrop.Key == 1 || (oxcrop.Key == 2 && this.isSecondSUTServerConnected))
                        {
                            this.HardDeleteMessageAndSubfolders(oxcrop.Key);
                        }
                    }
                }

                foreach (KeyValuePair<int, OxcropsClient> oxcrop in this.oxcropsClient)
                {
                    bool result = this.oxcropsClient[oxcrop.Key].Disconnect();
                    string comment = "Disconnecting server" + oxcrop.Key + " should be successful.";
                    Site.Assert.IsTrue(result, comment);
                }
            }
            catch (SEHException e)
            {
                Site.Assume.Fail(e.Message);
            }
            finally
            {
                this.isImportMessageMoveROP = false;
                this.importMessageTimes = 0;
                this.isNonImportMessageChangeOperation = false;
                this.lastChangeMadeByServer = false;
                this.lastChangeMadeByClient = false;
                this.importFlag = ImportFlag.InvalidParameter;
                this.importPidTagChangeKeyValue = new byte[22];
                this.previousGetBufferResult = RopResult.Success;
                this.needDoCleanup = true;
                this.isOrderByDeliveryTimeExtraFlagNotSet = false;
                AdapterHelper.ClearIndex();
                this.isSecondSUTServerConnected = false;
                this.localIdOffSet = 0;
                this.localId = new byte[6];
                this.serverReplicaGuid = new Guid();
                this.localReplicaGuid = Guid.NewGuid();
                this.handleContainer.Clear();
                this.objectIdContainer.Clear();
                this.streamBufferContainer.Clear();
                this.icsStateContainer.Clear();
                this.existNoPermissionFolder = false;
                this.cnsetRead.Clear();
                this.cnsetSeen.Clear();
                this.cnsetSeenFAI.Clear();
                base.Reset();
            }
        }

        /// <summary>
        /// Connect to the server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="connectionType">The type of connection</param>
        public void Connect(int serverId, ConnectionType connectionType)
        {
            string tempDomainValue = Common.GetConfigurationPropertyValue(Domain, this.Site);
            string tempUserNameValue = string.Empty;
            string tempPasswordValue = string.Empty;
            string tempServer = string.Empty;
            string tempUserDNValue = string.Empty;
            if (serverId == 1)
            {
                tempUserNameValue = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
                tempPasswordValue = Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site);
                tempServer = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
                tempUserDNValue = Common.GetConfigurationPropertyValue("AdminUserESSDN", this.Site) + "\0";
            }

            if (serverId == 2)
            {
                tempUserNameValue = Common.GetConfigurationPropertyValue("User2Name", this.Site);
                tempPasswordValue = Common.GetConfigurationPropertyValue("User2Password", this.Site);
                tempServer = Common.GetConfigurationPropertyValue("Sut2ComputerName", this.Site);
                tempUserDNValue = Common.GetConfigurationPropertyValue("TestUser2ESSDN", this.Site) + "\0";
            }

            bool ret = this.oxcropsClient[serverId].Connect(tempServer, connectionType, tempUserDNValue, tempDomainValue, tempUserNameValue, tempPasswordValue);

            if (serverId == 2)
            {
                this.isSecondSUTServerConnected = true;
            }

            Site.Assert.IsTrue(ret, "Connecting to server {0} should be successful", serverId);
        }

        /// <summary>
        /// Disconnect the connection to server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        public void Disconnect(int serverId)
        {
            bool ret = this.oxcropsClient[serverId].Disconnect();
            Site.Assert.IsTrue(ret, "Disconnecting server {0} should be successful", serverId);
        }

        /// <summary>
        /// Release the object by handle.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The object handle index.</param>
        /// <returns>The ROP result.</returns>
        public RopResult Release(int serverId, int objHandleIndex)
        {
            // If objHandleIndex is equal to -1, it indicates that the exchange version is 2010
            if (objHandleIndex == -1)
            {
                return RopResult.Success;
            }

            uint objHandle = this.handleContainer[objHandleIndex];

            RopReleaseRequest releaseRequest;
            releaseRequest.RopId = 0x01;
            releaseRequest.LogonId = 0x00;
            releaseRequest.InputHandleIndex = 0x00;
            this.Process(serverId, releaseRequest, objHandle);
            return RopResult.Success;
        }

        /// <summary>
        /// Logon the server.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="flag">The type of logon.</param>
        /// <param name="logonHandleIndex">The server object handle index.</param>
        /// <param name="inboxFolderIdIndex">The inbox folder Id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult Logon(int serverId, LogonFlags flag, out int logonHandleIndex, out int inboxFolderIdIndex)
        {
            // Initialize  return value.
            logonHandleIndex = -1;
            inboxFolderIdIndex = -1;
            uint inputHandle = 0;
            uint outputHandle = 0;

            // Form ROP request
            RopLogonRequest logonRequest;
            logonRequest.RopId = 0xFE;
            logonRequest.LogonId = 0x00;
            logonRequest.OutputHandleIndex = 0x00;
            logonRequest.StoreState = 0;
            if (LogonFlags.PublicFolder == flag || LogonFlags.Ghosted == flag)
            {
                logonRequest.LogonFlags = (byte)LogonFlags.PublicFolder;

                // Set PUBLIC and USE_PER_MDB_REPLID_MAPPING flag
                logonRequest.OpenFlags = 0x01000002;
                logonRequest.EssdnSize = 0;
                logonRequest.Essdn = new byte[0];
            }
            else
            {
                logonRequest.LogonFlags = (byte)LogonFlags.Private;

                // USE_PER_MDB_REPLID_MAPPING flag
                logonRequest.OpenFlags = 0x01000000;
                string tempUserDNValue = string.Empty;
                if (serverId == 1)
                {
                    tempUserDNValue = Common.GetConfigurationPropertyValue("AdminUserESSDN", this.Site) + "\0";
                }

                if (serverId == 2)
                {
                    tempUserDNValue = Common.GetConfigurationPropertyValue("TestUser2ESSDN", this.Site) + "\0";
                }

                logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(tempUserDNValue);
                logonRequest.Essdn = Encoding.ASCII.GetBytes(tempUserDNValue);
            }

            // Set the PutBuffer operation's default data
            RopLogonResponse logonResponse = (RopLogonResponse)this.Process(serverId, logonRequest, inputHandle);

            // Get return value.
            RopResult result = (RopResult)logonResponse.ReturnValue;
            string folderName;
            switch (flag)
            {
                case LogonFlags.Private:
                    if (result == RopResult.Success)
                    {
                        outputHandle = this.responseSOHs[logonResponse.OutputHandleIndex];
                        this.logonHandleOut = this.responseSOHs[logonResponse.OutputHandleIndex];
                        logonHandleIndex = AdapterHelper.GetHandleIndex();
                        this.handleContainer.Add(logonHandleIndex, outputHandle);
                        this.inboxFolderId = logonResponse.FolderIds[4];
                        inboxFolderIdIndex = AdapterHelper.GetObjectIdIndex();
                        this.objectIdContainer.Add(inboxFolderIdIndex, this.inboxFolderId);
                        return result;
                    }
                    else
                    {
                        logonHandleIndex = -1;
                        inboxFolderIdIndex = -1;
                        return result;
                    }

                case LogonFlags.PublicFolder:
                    folderName = Common.GetConfigurationPropertyValue("PublicFolderName", this.Site);
                    break;
                case LogonFlags.Ghosted:
                    folderName = Common.GetConfigurationPropertyValue("GhostedPublicFolderName", this.Site);
                    break;
                default:
                    AdapterHelper.Site.Assert.Fail("The given LogonFlag is unknown, its value is {0}.", flag);
                    folderName = string.Empty;
                    break;
            }
            
            // The code below is only for public folder logon
            if (result == RopResult.Success)
            {
                outputHandle = this.responseSOHs[logonResponse.OutputHandleIndex];
                this.logonHandleOut = this.responseSOHs[logonResponse.OutputHandleIndex];
                logonHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(logonHandleIndex, outputHandle);
                this.inboxFolderId = logonResponse.FolderIds[1];
                this.rootFolderId = this.inboxFolderId;

                ulong folderID;
                uint folderHandle;
                this.GetPublicFolderHandle(
                    serverId,
                    this.inboxFolderId,
                    folderName,
                    out folderID,
                    out folderHandle);
                this.publicFolderId = folderID;
                inboxFolderIdIndex = AdapterHelper.GetObjectIdIndex();
                this.objectIdContainer.Add(inboxFolderIdIndex, folderID);
                return result;
            }
            else
            {
                logonHandleIndex = -1;
                inboxFolderIdIndex = -1;
                return result;
            }
        }

        /// <summary>
        /// Open specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="folderHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult OpenFolder(int serverId, int objHandleIndex, int folderIdIndex, out int folderHandleIndex)
        {
            // Initialize return value.
            RopResult result = RopResult.InvalidParameter;
            folderHandleIndex = -1;

            // Form ROP request.
            uint objHandle = this.handleContainer[objHandleIndex];
            ulong folderId = this.objectIdContainer[folderIdIndex];
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = 0x02,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderId = folderId,

                // ReadOnly
                OpenModeFlags = 0x00
            };

            // Send request and get response.
            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.Process(serverId, openFolderRequest, objHandle);
            result = (RopResult)openFolderResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                folderHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(folderHandleIndex, this.responseSOHs[openFolderResponse.OutputHandleIndex]);
            }

            return result;
        }

        /// <summary>
        /// Open a specific message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The handle index folder object which the message in. </param>
        /// <param name="folderIdIndex">The folder id index of which the specific message in.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <param name="openMessageHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult OpenMessage(int serverId, int folderHandleIndex, int folderIdIndex, int messageIdIndex, out int openMessageHandleIndex)
        {
            // Initialize return value.
            RopResult result = RopResult.InvalidParameter;
            openMessageHandleIndex = -1;

            // Form ROP request.
            RopOpenMessageRequest openMessageRequest;
            uint openMesageHandle = this.handleContainer[folderHandleIndex];
            ulong folderId = this.objectIdContainer[folderIdIndex];
            ulong messageId = this.objectIdContainer[messageIdIndex];
            openMessageRequest.RopId = 0x03;
            openMessageRequest.LogonId = 0x00;

            // Set InputHandleIndex to 0x0, which specifies the location in the server object handle table where the handle,
            openMessageRequest.InputHandleIndex = 0x00;

            // Set OutputHandleIndex to 0x01, which specifies the location in the server object handle table where the handle,
            openMessageRequest.OutputHandleIndex = 0x01;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used,
            openMessageRequest.CodePageId = 0x0fff;

            // Set FolderId to the 5th of logonResponse(INBOX), which identifies the parent folder,
            openMessageRequest.FolderId = folderId;

            // Message will be opened for both reading and writing.
            openMessageRequest.OpenModeFlags = 0X01;

            // Set MessageId to that of created message, which identifies the message to be opened,
            openMessageRequest.MessageId = messageId;

            // Send request and get response.
            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.Process(serverId, openMessageRequest, openMesageHandle);
            result = (RopResult)openMessageResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                openMessageHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(openMessageHandleIndex, this.responseSOHs[openMessageResponse.OutputHandleIndex]);
                result = RopResult.Success;
            }

            return result;
        }

        /// <summary>
        /// Retrieve the content table for a folder. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult GetContentsTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount)
        {
            RopResult result = RopResult.InvalidParameter;
            rowCount = -1;
            uint objHandle = this.handleContainer[folderHandleIndex];

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse = new RopGetContentsTableResponse();

            getContentsTableRequest.RopId = 0x05;
            getContentsTableRequest.LogonId = 0x00;
            getContentsTableRequest.InputHandleIndex = 0x00;
            getContentsTableRequest.OutputHandleIndex = 0x01;

            // If this bit is set, the contents table lists only the soft-deleted messages. 
            getContentsTableRequest.TableFlags = 0x20;

            getContentsTableResponse = (RopGetContentsTableResponse)this.Process(serverId, getContentsTableRequest, objHandle);
            result = (RopResult)getContentsTableResponse.ReturnValue;
            int expectedRowCount = (int)getContentsTableResponse.RowCount;

            // Add this condition to match the model logical for return result, it's better to compare.
            if (deleteFlags == DeleteFlags.SoftDeleteCheck && this.currentSoftDeleteRowCount + 1 == expectedRowCount)
            {
                rowCount = 1;
            }
            else if (deleteFlags == DeleteFlags.HardDeleteCheck && this.currentSoftDeleteRowCount == expectedRowCount)
            {
                rowCount = 0;
            }
            else if (deleteFlags == DeleteFlags.Initial)
            {
                rowCount = 0;
            }
            else if (deleteFlags == DeleteFlags.SoftDeleteCheck && this.isImportMessageMoveROP && this.currentSoftDeleteRowCount + 1 == expectedRowCount)
            {
                rowCount = 1;
            }

            this.currentSoftDeleteRowCount = expectedRowCount;
            return result;
        }

        /// <summary>
        /// Retrieve the hierarchy table for a folder. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index</param>
        /// <param name="deleteFlags">The delete flag indicates whether checking delete.</param>
        /// <param name="rowCount">The row count.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult GetHierarchyTable(int serverId, int folderHandleIndex, DeleteFlags deleteFlags, out int rowCount)
        {
            RopResult result = RopResult.InvalidParameter;
            rowCount = -1;
            uint objHandle = this.handleContainer[folderHandleIndex];

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest();
            RopGetHierarchyTableResponse getHierarchyTableResponse = new RopGetHierarchyTableResponse();

            getHierarchyTableRequest.RopId = 0x04;
            getHierarchyTableRequest.LogonId = 0x00;
            getHierarchyTableRequest.InputHandleIndex = 0x00;
            getHierarchyTableRequest.OutputHandleIndex = 0x01;

            // If this bit is set, the contents table lists only the soft-deleted messages. 
            getHierarchyTableRequest.TableFlags = 0x20;

            getHierarchyTableResponse = (RopGetHierarchyTableResponse)this.Process(serverId, getHierarchyTableRequest, objHandle);
            result = (RopResult)getHierarchyTableResponse.ReturnValue;
            int expectedRowCount = (int)getHierarchyTableResponse.RowCount;

            // Add this condition to match the model logical for return result, it's better to compare.
            if (deleteFlags == DeleteFlags.SoftDeleteCheck && this.currentSoftDeleteRowCount + 1 == expectedRowCount)
            {
                rowCount = 1;
            }
            else if (deleteFlags == DeleteFlags.HardDeleteCheck && this.currentSoftDeleteRowCount == expectedRowCount && Common.IsRequirementEnabled(90205002, this.Site))
            {
                rowCount = 0;
            }
            else if (deleteFlags == DeleteFlags.Initial)
            {
                rowCount = 0;
            }

            this.currentSoftDeleteRowCount = expectedRowCount;
            return result;
        }

        /// <summary>
        /// Create a message and return the message handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index for creating message.</param>
        /// <param name="folderIdIndex">The folder Id index.</param>
        /// <param name="associatedFlag">The message is FAI or not.</param>
        /// <param name="messageHandleIndex">The created message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult CreateMessage(int serverId, int folderHandleIndex, int folderIdIndex, bool associatedFlag, out int messageHandleIndex)
        {
            // Initialize operation values.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse = new RopCreateMessageResponse();
            messageHandleIndex = -1;
            uint folderHandle = this.handleContainer[folderHandleIndex];
            ulong folderId = this.objectIdContainer[folderIdIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct the RopCreateMessage request.
            createMessageRequest.RopId = 0x06;
            createMessageRequest.LogonId = 0x00;
            createMessageRequest.InputHandleIndex = 0x00;
            createMessageRequest.OutputHandleIndex = 0x01;

            // Set CodePageId to 0x0FFF, which specified the code page of Logon object will be used
            createMessageRequest.CodePageId = 0x0FFF;
            createMessageRequest.FolderId = folderId;

            // If the message is a FAImessage.
            createMessageRequest.AssociatedFlag = associatedFlag ? (byte)0x01 : (byte)0x00;

            // Send request and get ROP response.
            createMessageResponse = (RopCreateMessageResponse)this.Process(serverId, createMessageRequest, folderHandle);
            result = (RopResult)createMessageResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                messageHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(messageHandleIndex, this.responseSOHs[createMessageResponse.OutputHandleIndex]);

                this.AddRecipient(serverId, this.responseSOHs[createMessageResponse.OutputHandleIndex]);

                // Verify server Accept PidTagIdsetGivenPtypBinary
                this.VerifyServerAcceptPidTagIdsetGivenPtypBinary(serverId, folderHandleIndex);
            }

            return result;
        }

        /// <summary>
        /// Create an attachment on specific message object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="messageHandleIndex">The message handle</param>
        /// <param name="attachmentHandleIndex">The attachment handle of created</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult CreateAttachment(int serverId, int messageHandleIndex, out int attachmentHandleIndex)
        {
            // Initialize ROP operation data.
            RopCreateAttachmentRequest createAttachmentRequest;
            attachmentHandleIndex = -1;
            RopResult result = RopResult.InvalidParameter;
            uint handle = this.handleContainer[messageHandleIndex];

            // Construct ROP request.
            createAttachmentRequest.RopId = 0x23;
            createAttachmentRequest.LogonId = 0x00;
            createAttachmentRequest.InputHandleIndex = 0x00;
            createAttachmentRequest.OutputHandleIndex = 0x01;

            // Send request and get response.
            RopCreateAttachmentResponse createAttachmentResponse = (RopCreateAttachmentResponse)this.Process(serverId, createAttachmentRequest, handle);
            result = (RopResult)createAttachmentResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                attachmentHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(attachmentHandleIndex, this.responseSOHs[createAttachmentResponse.OutputHandleIndex]);
            }

            return result;
        }

        /// <summary>
        /// Save the changes property of message.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="messageHandleIndex">The message handle index.</param>
        /// <param name="messageIdIndex">The message id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SaveChangesMessage(int serverId, int messageHandleIndex, out int messageIdIndex)
        {
            // Initialize ROP operation data.
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            messageIdIndex = -1;
            RopResult result = RopResult.InvalidParameter;
            uint saveChangeMessageHandle = this.handleContainer[messageHandleIndex];

            // Construct the RopSaveChangesMessage request.
            saveChangesMessageRequest.RopId = 0x0c;
            saveChangesMessageRequest.LogonId = 0x00;
            saveChangesMessageRequest.InputHandleIndex = 0x00;
            saveChangesMessageRequest.ResponseHandleIndex = 0x01;

            // Default with no specifications.
            saveChangesMessageRequest.SaveFlags = 0x0C;

            // Send the RopSaveChangesMessage request 
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.Process(serverId, saveChangesMessageRequest, saveChangeMessageHandle);
            result = (RopResult)saveChangesMessageResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                if (this.importFlag != ImportFlag.InvalidParameter && !this.lastChangeMadeByServer)
                {
                    this.VerifyInternalIdentifier(saveChangesMessageResponse.MessageId);
                }

                if (this.objectIdContainer.ContainsValue((ulong)saveChangesMessageResponse.MessageId))
                {
                    foreach (int messageKey in this.objectIdContainer.Keys)
                    {
                        if (this.objectIdContainer[messageKey] == (ulong)saveChangesMessageResponse.MessageId)
                        {
                            messageIdIndex = messageKey;
                            break;
                        }
                    }
                }
                else
                {
                    messageIdIndex = AdapterHelper.GetObjectIdIndex();
                    this.objectIdContainer.Add(messageIdIndex, (ulong)saveChangesMessageResponse.MessageId);
                }

                result = RopResult.Success;

                // Verify Whether PCL Equal PidTagChangeKey
                this.VerifyValuePidTagPredecessorChangeListAndPidTagChangeKey(serverId, messageHandleIndex);
            }

            return result;
        }

        /// <summary>
        /// Commits the changes made to the Attachment object.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="attachmentHandleIndex">The attachment handle</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SaveChangesAttachment(int serverId, int attachmentHandleIndex)
        {
            // Initialize ROP operation data.
            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest;
            uint handle = this.handleContainer[attachmentHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            saveChangesAttachmentRequest.RopId = 0x25;
            saveChangesAttachmentRequest.LogonId = 0x00;
            saveChangesAttachmentRequest.ResponseHandleIndex = 0x01;
            saveChangesAttachmentRequest.InputHandleIndex = 0x00;

            // KeepOpenReadWrite,specified 
            saveChangesAttachmentRequest.SaveFlags = 0x0A;

            // Send the RopSaveChangesMessage request 
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse = (RopSaveChangesAttachmentResponse)this.Process(serverId, saveChangesAttachmentRequest, handle);
            result = (RopResult)saveChangesAttachmentResponse.ReturnValue;

            return result;
        }

        /// <summary>
        /// Create a folder and return the folder handle created.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="folderName">The new folder's name.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <param name="folderHandleIndex">The new folder's handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult CreateFolder(int serverId, int objHandleIndex, string folderName, out int folderIdIndex, out int folderHandleIndex)
        {
            // Initialize return value.
            folderIdIndex = -1;
            folderHandleIndex = -1;
            uint objHandle = this.handleContainer[objHandleIndex];
            RopResult result = RopResult.InvalidParameter;
            byte[] getName = Encoding.ASCII.GetBytes(folderName + AdapterHelper.GetFolderNameIndex() + "\0");

            // Construct ROP request.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = 0x1C,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderType = 0x01,
                UseUnicodeStrings = 0x00,
                OpenExisting = 0x01,
                Reserved = 0x00,
                DisplayName = getName,
                Comment = getName
            };

            // Send request and get response data.
            RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.Process(serverId, createFolderRequest, objHandle);
            result = (RopResult)createFolderResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                folderIdIndex = AdapterHelper.GetObjectIdIndex();
                this.objectIdContainer.Add(folderIdIndex, (ulong)createFolderResponse.FolderId);
                folderHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(folderHandleIndex, this.responseSOHs[createFolderResponse.OutputHandleIndex]);
                result = RopResult.Success;
                return result;
            }

            return result;
        }

        /// <summary>
        /// Create a folder and return the folder handle for the newly created folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandle">The server object handle</param>
        /// <param name="folderName">The new folder's name.</param>
        /// <param name="folderIdIndex">The folder id index</param>
        /// <param name="folderHandleIndex">The new folder's handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult CreateFolder(int serverId, uint objHandle, string folderName, out int folderIdIndex, out int folderHandleIndex)
        {
            // Initialize return value.
            folderIdIndex = -1;
            folderHandleIndex = -1;
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = 0x1C,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderType = 0x01,
                UseUnicodeStrings = 0x00,
                OpenExisting = 0x01,
                Reserved = 0x00,
                DisplayName = Encoding.ASCII.GetBytes(folderName + "\0"),
                Comment = Encoding.ASCII.GetBytes(folderName + "\0")
            };

            // Send request and get response data.
            RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.Process(serverId, createFolderRequest, objHandle);
            result = (RopResult)createFolderResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                folderIdIndex = AdapterHelper.GetObjectIdIndex();
                this.objectIdContainer.Add(folderIdIndex, (ulong)createFolderResponse.FolderId);
                folderHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(folderHandleIndex, this.responseSOHs[createFolderResponse.OutputHandleIndex]);
            }

            return result;
        }

        /// <summary>
        /// Delete the specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The folder handle index.</param>
        /// <param name="folderIdIndex">The folder id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult DeleteFolder(int serverId, int folderHandleIndex, int folderIdIndex)
        {
            // Initialize ROP operation data.
            uint folderHandle = this.handleContainer[folderHandleIndex];
            ulong folderId = this.objectIdContainer[folderIdIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = 0x1D,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                DeleteFolderFlags = 0x15,
                FolderId = folderId
            };

            // DEL_MESSAGES, DEL_FOLDERS and DELETE_HARD_DELETE.

            // Send request and get response.
            RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.Process(serverId, deleteFolderRequest, folderHandle);
            result = (RopResult)deleteFolderResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                this.objectIdContainer.Remove(folderIdIndex);
            }

            return result;
        }

        /// <summary>
        /// Delete the specific folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandle">The folder handle</param>
        /// <param name="folderId">The folder id</param>
        /// <returns>Indicate the result of this ROP operation</returns>
        public RopResult DeleteFolder(int serverId, uint folderHandle, ulong folderId)
        {
            // Initialize ROP operation data.
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = 0x1D,
                LogonId = 0x00,
                InputHandleIndex = 0x00,

                // DEL_MESSAGES, DEL_FOLDERS and DELETE_HARD_DELETE.
                DeleteFolderFlags = 0x15,
                FolderId = folderId
            };

            // Send request and get response.
            RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.Process(serverId, deleteFolderRequest, folderHandle);
            result = (RopResult)deleteFolderResponse.ReturnValue;
            return result;
        }

        #endregion ROP Interfaces implement

        #region IMS_OXCFXICSAdapter Members

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant subobjects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">Type of object handle </param>
        /// <param name="level">Variable indicate whether copy the descendant subobjects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="propertyTags">Array of properties and subobjects to exclude.</param>
        /// <param name="copyToHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferSourceCopyTo(int serverId, int sourceHandleIndex, InputHandleType handleType, bool level, CopyToCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int copyToHandleIndex)
        {
            // Initialize ROP returned value.
            RopResult result = RopResult.InvalidParameter;
            copyToHandleIndex = -1;
            ushort propertyCount = (ushort)propertyTags.Count;
            this.currentCopyFlag = copyFlag;

            // Get stream type.
            switch (handleType)
            {
                case InputHandleType.FolderHandle:
                    this.streamType = FastTransferStreamType.folderContent;
                    break;
                case InputHandleType.MessageHandle:
                    this.streamType = FastTransferStreamType.MessageContent;
                    break;
                case InputHandleType.AttachmentHandle:
                    this.streamType = FastTransferStreamType.attachmentContent;
                    break;
                default:
                    break;
            }

            // Initialize input propertyTags
            PropertyTag[] messagePropertyTags = new PropertyTag[propertyCount];
            for (int i = 0; i < propertyCount; i++)
            {
                messagePropertyTags[i] = this.propertyTagsDictionary[propertyTags[i]];
            }

            // Stores the values for further verification
            this.previousOperation = EnumFastTransferOperation.FastTransferSourceCopyTo;
            this.propertyTags = messagePropertyTags;

            // Construct ROP request.
            uint sourceHandle = this.handleContainer[sourceHandleIndex];
            RopFastTransferSourceCopyToRequest request;
            request.RopId = 0x4D;
            request.LogonId = 0X00;
            request.InputHandleIndex = 0x00;
            request.OutputHandleIndex = 0x01;

            // This value specifies the level at which the copy is occurring,Non-Zero: exclude all descendant subobjects from being copied
            request.Level = level ? (byte)1 : (byte)0;
            request.CopyFlags = (uint)copyFlag;
            request.SendOptions = (byte)option;
            request.PropertyTagCount = propertyCount;
            request.PropertyTags = messagePropertyTags;

            // Send request and get response.
            RopFastTransferSourceCopyToResponse response = (RopFastTransferSourceCopyToResponse)this.Process(serverId, request, sourceHandle);
            result = (RopResult)response.ReturnValue;

            if (result == RopResult.Success)
            {
                copyToHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(copyToHandleIndex, this.responseSOHs[response.OutputHandleIndex]);
            }

            // Verify ROP FastTransferSourceCopyTo
            this.VerifyRopFastTransferSourceCopyTo(request, response);

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation to download content from a given messaging object and its descendant subobjects.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder or message object handle index. </param>
        /// <param name="handleType">Type of object handle. </param>
        /// <param name="level">Variable indicate whether copy the descendant subobjects.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="propertyTags">The list of properties and subobjects to exclude.</param>
        /// <param name="copyPropertiesHandleIndex">The properties handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferSourceCopyProperties(int serverId, int sourceHandleIndex, InputHandleType handleType, bool level, CopyPropertiesCopyFlags copyFlag, SendOptionAlls option, Sequence<string> propertyTags, out int copyPropertiesHandleIndex)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;
            RopFastTransferSourceCopyPropertiesRequest req;
            copyPropertiesHandleIndex = -1;
            uint sourceHandle = this.handleContainer[sourceHandleIndex];

            // Get stream type.
            switch (handleType)
            {
                case InputHandleType.FolderHandle:
                    this.streamType = FastTransferStreamType.folderContent;
                    break;
                case InputHandleType.MessageHandle:
                    this.streamType = FastTransferStreamType.MessageContent;
                    break;
                case InputHandleType.AttachmentHandle:
                    this.streamType = FastTransferStreamType.attachmentContent;
                    break;
                default:
                    break;
            }

            // Initialize input propertyTags.
            ushort propertyCount = (ushort)propertyTags.Count;
            PropertyTag[] messageSamplePropertyTags = new PropertyTag[propertyCount];
            for (int i = 0; i < propertyCount; i++)
            {
                messageSamplePropertyTags[i] = this.propertyTagsDictionary[propertyTags[i]];
            }

            this.previousOperation = EnumFastTransferOperation.FastTransferSourceCopyProperties;
            this.propertyTags = messageSamplePropertyTags;
            this.copySubObjects = !level;

            // Construct ROP request.
            req.RopId = 0x69;
            req.LogonId = 0x00;
            req.InputHandleIndex = 0x00;
            req.OutputHandleIndex = 0x01;

            // This value specifies the level at which the copy is occurring, which is specified in [MS-OXCROPS].
            req.Level = (byte)(level ? 1 : 0);

            req.CopyFlags = (byte)copyFlag;
            req.SendOptions = (byte)option;
            req.PropertyTagCount = propertyCount;
            req.PropertyTags = messageSamplePropertyTags;

            // Send request and get response.
            RopFastTransferSourceCopyPropertiesResponse response = (RopFastTransferSourceCopyPropertiesResponse)this.Process(serverId, req, sourceHandle);
            result = (RopResult)response.ReturnValue;

            if (result == RopResult.Success)
            {
                copyPropertiesHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(copyPropertiesHandleIndex, this.responseSOHs[response.OutputHandleIndex]);
            }

            // Verify ROP FastTransferSourceCopyProperties
            this.VerifyRopFastTransferSourceCopyProperties(req, response);

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation on a folder for downloading content and descendant subobjects for messages identified by a given set of IDs.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="messageIds">The list of MIDs the messages should copy.</param>
        /// <param name="copyMessageHandleIndex">The message handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferSourceCopyMessages(int serverId, int sourceHandleIndex, RopFastTransferSourceCopyMessagesCopyFlags copyFlag, SendOptionAlls option, Sequence<int> messageIds, out int copyMessageHandleIndex)
        {
            // Initialize ROP data
            RopResult result = RopResult.InvalidParameter;
            copyMessageHandleIndex = -1;
            this.streamType = FastTransferStreamType.MessageList;
            RopFastTransferSourceCopyMessagesRequest req;
            uint sourceHandle = this.handleContainer[sourceHandleIndex];
            ushort idcount = (ushort)messageIds.Count;
            ulong[] messageId = new ulong[idcount];
            int index = 0;
            foreach (int mID in messageIds)
            {
                messageId[index++] = this.objectIdContainer[mID];
            }

            // Construct ROP request.
            req.RopId = 0x4B;
            req.LogonId = 0x00;
            req.InputHandleIndex = 0x00;
            req.OutputHandleIndex = 0x01;
            req.MessageIdCount = idcount;
            req.MessageIds = messageId;
            req.CopyFlags = (byte)copyFlag;
            req.SendOptions = (byte)option;

            // Stores message id and copy flag
            this.previousOperation = EnumFastTransferOperation.FastTransferSourceCopyMessage;
            this.messageIdForFastTransferSourceCopyMessages = (long)messageId[0];
            this.copyFlagForFastTransferSourceCopyMessages = copyFlag;

            // Send request and get response.
            RopFastTransferSourceCopyMessagesResponse response = (RopFastTransferSourceCopyMessagesResponse)this.Process(serverId, req, sourceHandle);
            result = (RopResult)response.ReturnValue;

            if (result == RopResult.Success)
            {
                copyMessageHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(copyMessageHandleIndex, this.responseSOHs[response.OutputHandleIndex]);
            }

            // Verify ROP FastTransferSourceCopyMessages
            this.VerifyRopFastTransferSourceCopyMessages(req, response);

            return result;
        }

        /// <summary>
        /// Initializes a FastTransfer operation to download properties and descendant subobjects for a specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Folder object handle index. </param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="copyFolderHandleIndex">The folder handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferSourceCopyFolder(int serverId, int sourceHandleIndex, CopyFolderCopyFlags copyFlag, SendOptionAlls option, out int copyFolderHandleIndex)
        {
            // Initialize return value.
            copyFolderHandleIndex = -1;
            RopResult result = RopResult.InvalidParameter;
            uint targetFolderHandle = this.handleContainer[sourceHandleIndex];

            // Move flag only used in exchangeServer2007
            if (!Common.IsRequirementEnabled(526001, this.Site) && copyFlag == CopyFolderCopyFlags.Move)
            {
                return result;
            }

            this.streamType = FastTransferStreamType.TopFolder;

            // Construct ROP request.
            RopFastTransferSourceCopyFolderRequest fastTransferSourceCopyFolderRequest;
            fastTransferSourceCopyFolderRequest.RopId = 0x4C;
            fastTransferSourceCopyFolderRequest.LogonId = 0x00;
            fastTransferSourceCopyFolderRequest.InputHandleIndex = 0x00;
            fastTransferSourceCopyFolderRequest.OutputHandleIndex = 0x01;
            fastTransferSourceCopyFolderRequest.CopyFlags = (byte)copyFlag;
            fastTransferSourceCopyFolderRequest.SendOptions = (byte)option;

            this.folderCopyFlag = copyFlag;

            // Send request and get response.
            RopFastTransferSourceCopyFolderResponse response = (RopFastTransferSourceCopyFolderResponse)this.Process(serverId, fastTransferSourceCopyFolderRequest, targetFolderHandle);
            result = (RopResult)response.ReturnValue;

            if (result == RopResult.Success)
            {
                copyFolderHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(copyFolderHandleIndex, this.responseSOHs[response.OutputHandleIndex]);
            }

            // Verify ROP FastTransferSourceCopyFolder
            this.VerifyRopFastTransferSourceCopyFolder(fastTransferSourceCopyFolderRequest, response);

            return result;
        }

        /// <summary>
        /// Tell the server of another server's version.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">Server object handle index in handle container.</param>
        /// <param name="otherServerId">Another server's id.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult TellVersion(int serverId, int sourceHandleIndex, int otherServerId)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;

            if (sourceHandleIndex < 0)
            {
                return result;
            }

            uint inobjHandle = this.handleContainer[sourceHandleIndex];

            // Construct ROP request.
            RopTellVersionRequest tellVersionRequest;

            ushort[] version = this.oxcropsClient[serverId].MapiContext.EXServerVersion;
            tellVersionRequest.RopId = 0x86;
            tellVersionRequest.LogonId = 0x00;
            tellVersionRequest.InputHandleIndex = 0x00;
            tellVersionRequest.Version = version;

            // Send request and get response data.
            RopTellVersionResponse response = (RopTellVersionResponse)this.Process(serverId, tellVersionRequest, inobjHandle);
            result = (RopResult)response.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP TellVersion
                this.VerifyRopTellVersion(response);
            }

            return result;
        }

        /// <summary>
        /// Downloads the next portion of a FastTransfer stream.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="downloadHandleIndex">A fast transfer stream object handle index. </param>
        /// <param name="bufferSize">Specifies the maximum amount of data to be output in the TransferBuffer.</param>
        /// <param name="transferBufferIndex">The index of data get from the fast transfer stream.</param>
        /// <param name="abstractFastTransferStream">Fast transfer stream.</param>
        /// <param name="transferDataSmallOrEqualToBufferSize">Variable to not if the transferData is small or equal to bufferSize</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferSourceGetBuffer(int serverId, int downloadHandleIndex, BufferSize bufferSize, out int transferBufferIndex, out AbstractFastTransferStream abstractFastTransferStream, out bool transferDataSmallOrEqualToBufferSize)
        {
            // Initialize ROP data.
            SyntacticalBase.AllPropList = null;
            RopResult returnValue = RopResult.InvalidParameter;
            this.totalTransferBufferList.Clear();
            transferBufferIndex = -1;
            abstractFastTransferStream = new AbstractFastTransferStream();
            transferDataSmallOrEqualToBufferSize = false;
            if (downloadHandleIndex < 0)
            {
                return returnValue;
            }

            if (bufferSize == BufferSize.Greater && !Common.IsRequirementEnabled(2625, this.Site))
            {
                returnValue = RopResult.BufferTooSmall;
                return returnValue;
            }

            RopFastTransferSourceGetBufferResponse response = new RopFastTransferSourceGetBufferResponse();
            uint sourceHandle = this.handleContainer[downloadHandleIndex];
            uint downloadContextHandle = sourceHandle;

            // Construct ROP request.
            RopFastTransferSourceGetBufferRequest fastTransferSourceGetBufferRequest;
            fastTransferSourceGetBufferRequest.RopId = 0x4e;
            fastTransferSourceGetBufferRequest.LogonId = 0x00;
            fastTransferSourceGetBufferRequest.InputHandleIndex = 0x00;
            fastTransferSourceGetBufferRequest.BufferSize = (ushort)bufferSize;
            if (bufferSize != BufferSize.Normal)
            {
                fastTransferSourceGetBufferRequest.MaximumBufferSize = null;
            }
            else
            {
                fastTransferSourceGetBufferRequest.MaximumBufferSize = (ushort)bufferSize;
            }

            bool isRunIntoPartial = false;
            bool isRunIntoNoRoom = false;
            do
            {
                IDeserializable tempRopResponse = null;
                if ((tempRopResponse = this.Process(serverId, fastTransferSourceGetBufferRequest, downloadContextHandle)) != null)
                {    // Send request and get response.
                    response = (RopFastTransferSourceGetBufferResponse)tempRopResponse;

                    byte[] transferBuffer = new byte[(int)response.TransferBufferSize];
                    if (response.ReturnValue == 0)
                    {
                        if (bufferSize != BufferSize.Normal)
                        {
                            transferDataSmallOrEqualToBufferSize = transferBuffer.Length <= (int)bufferSize;
                        }

                        for (int i = 0; i < (int)response.TransferBufferSize; i++)
                        {
                            transferBuffer[i] = response.TransferBuffer[i];
                        }

                        this.totalTransferBufferList.Add(transferBuffer);

                        if (response.TransferStatus == (ushort)TransferStatus.Partial)
                        {
                            isRunIntoPartial = true;
                        }

                        if (response.TransferStatus == (ushort)TransferStatus.NoRoom)
                        {
                            isRunIntoNoRoom = true;
                        }
                    }
                }
                else
                {
                    break;
                }
            }
            while (response.TransferStatus != (ushort)TransferStatus.Done && (RopResult)response.ReturnValue == RopResult.Success);

            if (response.TransferBuffer == null)
            {
                returnValue = (RopResult)this.ropResult;
            }
            else
            {
                returnValue = (RopResult)response.ReturnValue;
            }

            bool isReachedBufferTooSmall = (bufferSize == BufferSize.Greater) && (returnValue == RopResult.BufferTooSmall);

            this.VerifyTransferStatus(isRunIntoPartial, isRunIntoNoRoom, isReachedBufferTooSmall);

            if (isReachedBufferTooSmall)
            {
                this.previousGetBufferResult = RopResult.BufferTooSmall;
            }

            if (response.TransferStatus == (ushort)TransferStatus.Done)
            {
                returnValue = RopResult.Success;

                int bufferlength = 0;
                foreach (byte[] blengt in this.totalTransferBufferList)
                {
                    bufferlength += blengt.Length;
                }

                byte[] totalTransferBuffer = new byte[bufferlength];
                int index = 0;
                foreach (byte[] buffer in this.totalTransferBufferList)
                {
                    Array.Copy(buffer, 0, totalTransferBuffer, index, buffer.Length);
                    index += buffer.Length;
                }

                byte[] requiredTransferBuffer = this.ProcessFXSourceGetBuffer(totalTransferBuffer);

                using (FastTransferStream fs = new FastTransferStream(requiredTransferBuffer, true))
                {
                    // Verify FastTransfer Stream
                    this.VerifyFastTransferStream(fs, this.streamType);
                }

                abstractFastTransferStream = this.GenerateAbstractFastTransferStream(serverId, requiredTransferBuffer);
                if (bufferSize == BufferSize.Greater)
                {
                    transferBufferIndex = -1;
                }
                else
                {
                    transferBufferIndex = AdapterHelper.GetStreamBufferIndex();
                }

                int k = this.streamBufferContainer.Count;
                foreach (byte[] subBuffer in this.totalTransferBufferList)
                {
                    this.streamBufferContainer.Add(++k, subBuffer);
                }
            }

            if (response.RopId != 0x00)
            {
                // Verify ROP FastTransferSourceGetBuffer
                this.VerifyRopFastTransferSourceGetBuffer(response);
            }

            return returnValue;
        }

        /// <summary>
        /// Get content sync data.
        /// </summary>
        /// <param name="serverID">server id.</param>
        /// <param name="stateIndex">Start index.</param>
        /// <param name="cs">Content sync object</param>
        /// <returns>The AbstractFastTransferStream</returns>
        public AbstractFastTransferStream GetAbstractContentSync(
            int serverID,
            int stateIndex,
            ContentsSync cs)
        {
            AbstractFastTransferStream afts = new AbstractFastTransferStream();
            AbstractContentsSync acs = new AbstractContentsSync
            {
                AbstractDeletion = new AbstractDeletion
                {
                    IdSetDeleted = new Modeling.Set<int>()
                }
            };

            // If the deletion is not null change the AbstractDeletion's field,
            // else do not change the Deletions field of the AbstractContentsSync
            if (cs.Deletions != null)
            {
                // If the deletion is not null the deletion is present.
                acs.AbstractDeletion = new AbstractDeletion
                {
                    IsDeletionPresent = true,
                    IsPidTagIdsetExpiredExist = cs.Deletions.PropList.HasPidTagIdsetExpired,

                    // Check pidtags in the deletions property list.
                    IsPidTagIdsetNoLongerInScopeExist = cs.Deletions.PropList.HasPidTagIdsetNoLongerInScope,
                    IdSetDeleted = new Modeling.Set<int>()
                };

                // Check Idset in the deletion.
                // If the Idset that can be got from PidTagIdsetDeleted value 
                // in the PropList of the deletions is contained in the objectIdContainer
                // add it to the IdsetDeleted field in the AbstractDeletion
                List<ulong> ids = cs.Deletions.IdsetDeleted;
                foreach (int key in this.objectIdContainer.Keys)
                {
                    if (ids.Contains(this.objectIdContainer[key]))
                    {
                        acs.AbstractDeletion.IdSetDeleted =
                             acs.AbstractDeletion.IdSetDeleted.Add(key);
                    }
                }

                this.VerifyMetaTagIdsetDeletedProperty(this.hasExecuteSynchronizationImportDeletes, this.importDelFlag, ids, this.delObjId);
                this.hasExecuteSynchronizationImportDeletes = false;
            }

            // Whether ProgressTotal is present.
            acs.IsprogessTotalPresent = cs.ProgressTotal != null;

            // Whether ReadStateChanges is present.
            acs.IsReadStateChangesExist = cs.ReadStateChanges != null;

            acs.MessageInfo = new Modeling.Set<AbstractMessageChangeInfo>();
            if (cs.MessageChangeTuples != null)
            {
                for (int i = 0; i < cs.MessageChangeTuples.Count; i++)
                {
                    AbstractMessageChangeInfo info = new AbstractMessageChangeInfo();
                    Tuple<ProgressPerMessage, MessageChange> msg = cs.MessageChangeTuples[i];
                    info.IsMessageChangeFull = msg.Item2 is MessageChangeFull;

                    // Whether ProgressPerMessagePresent is present.
                    info.IsProgressPerMessagePresent = msg.Item1 != null;
                    if (msg.Item1 != null)
                    {
                        // If ProgressPerMessagePresent is present,
                        // whether the message followed is a FAI message.
                        info.FollowedFAIMessage = msg.Item1.IsFollowedFAIMessage;
                    }

                    info.IsPidTagChangeNumberExist = msg.Item2.HasPidTagChangeNumber;
                    info.IsPidTagMessageSizeExist = msg.Item2.HasPidTagMessageSize;
                    info.IsPidTagMidExist = msg.Item2.HasPidTagMid;

                    // convert the SourceKey in the messageChangeHeader of the messageChange
                    // to a GID  structure, then use this GID to build a LongTermId,
                    // get the id corresponding to the LongTermId,if the id is maintained in the 
                    // objectIdContainer,find its index, else set the messageIdIndex of the 
                    // AbstractMessageChangeInfo to -1.
                    byte[] buffer = msg.Item2.SourceKey;
                    GID gid = StructureSerializer.Deserialize<GID>(buffer);
                    LongTermId lid = new LongTermId
                    {
                        DatabaseGuid = gid.DatabaseGuid.ToByteArray(),
                        GlobalCounter = gid.GlobalCounter
                    };
                    ulong id = this.GetIdFromLongTerm(serverID, lid);
                    info.MessageIdIndex = -1;
                    foreach (int k in this.objectIdContainer.Keys)
                    {
                        if (this.objectIdContainer[k] == id)
                        {
                            info.MessageIdIndex = k;
                            break;
                        }
                    }

                    // If messageChangeFull, whether its PropList contains PidTagRtfCompressed. 
                    if (msg.Item2 is MessageChangeFull)
                    {
                        info.IsRTFformat = (msg.Item2 as MessageChangeFull).IsRTFFormat;
                    }

                    // Add the AbstractMessageChangeInfo to the MessageInfo of the AbstractContentsSync
                    acs.MessageInfo = acs.MessageInfo.Add(info);
                }

                // if there is  more than 1 messageChange presents.
                // whether their order is in descending order.
                if (cs.MessageChangeTuples.Count > 1)
                {
                    acs.IsSortByLastModificationTime = true;
                    acs.IsSortByMessageDeliveryTime = true;
                    DateTime prev = cs.MessageChangeTuples[0].Item2.LastModificationTime;
                    for (int i = 1; i < cs.MessageChangeTuples.Count; i++)
                    {
                        DateTime curr = cs.MessageChangeTuples[i].Item2.LastModificationTime;

                        // If one is not in descending order.
                        // set flags to false.
                        if (prev < curr)
                        {
                            acs.IsSortByLastModificationTime = false;
                            acs.IsSortByMessageDeliveryTime = false;
                            break;
                        }
                    }

                    // If the OrderByDeliveryTime flag of the SynchronizationExtraFlag field is not set, there is no requirement on the server to return items in a specific order.
                    if (this.isOrderByDeliveryTimeExtraFlagNotSet)
                    {
                        acs.IsSortByLastModificationTime = false;
                        acs.IsSortByMessageDeliveryTime = false;
                    }
                }
            }

            // Get FinalICSState.
            acs.FinalICSState = this.GetAbstractState(serverID, stateIndex, cs.State);
            afts.AbstractContentsSync = acs;
            afts.StreamType = FastTransferStreamType.contentsSync;
            afts.AbstractState = new AbstractState();
            return afts;
        }

        /// <summary>
        ///  Uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="transferDataIndex">Transfer data index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferDestinationPutBuffer(int serverId, int sourceHandleIndex, int transferDataIndex)
        {
            RopResult result = RopResult.Success;
            if (sourceHandleIndex < 0 || transferDataIndex <= 0)
            {
                return RopResult.InvalidParameter;
            }

            uint fastTransferDestinationPutBufferHandle = this.handleContainer[sourceHandleIndex];
            foreach (int i in this.streamBufferContainer.Keys)
            {
                byte[] transferData = this.streamBufferContainer[i];
                RopFastTransferDestinationPutBufferRequest fastTransferDestinationPutBufferRequest = new RopFastTransferDestinationPutBufferRequest();
                fastTransferDestinationPutBufferRequest.RopId = 0x54;
                fastTransferDestinationPutBufferRequest.LogonId = 0x00;
                fastTransferDestinationPutBufferRequest.InputHandleIndex = 0x00;
                fastTransferDestinationPutBufferRequest.TransferDataSize = (ushort)transferData.Length;
                fastTransferDestinationPutBufferRequest.TransferData = transferData;

                RopFastTransferDestinationPutBufferResponse response = (RopFastTransferDestinationPutBufferResponse)this.Process(serverId, fastTransferDestinationPutBufferRequest, fastTransferDestinationPutBufferHandle);

                if (RopResult.Success != (RopResult)response.ReturnValue)
                {
                    result = (RopResult)response.ReturnValue;
                    Site.Log.Add(LogEntryKind.Debug, "Call FastTransferDestinationPutBuffer fail at {0} times, return error: {1}", i, result);
                    break;
                }

                this.VerifyRopFastTransferDestinationPutBuffer(fastTransferDestinationPutBufferRequest, response);
            }

            return result;
        }

        /// <summary>
        ///  Uploads the next portion of an input FastTransfer stream for a previously configured FastTransfer upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="transferDataIndex">Transfer data index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferDestinationPutBufferExtended(int serverId, int sourceHandleIndex, int transferDataIndex)
        {
            RopResult result = RopResult.Success;
            if (sourceHandleIndex < 0 || transferDataIndex <= 0)
            {
                return RopResult.InvalidParameter;
            }

            uint fastTransferDestinationPutBufferExtendedHandle = this.handleContainer[sourceHandleIndex];
            foreach (int i in this.streamBufferContainer.Keys)
            {
                byte[] transferData = this.streamBufferContainer[i];
                RopFastTransferDestinationPutBufferExtendedRequest fastTransferDestinationPutBufferExtendedRequest = new RopFastTransferDestinationPutBufferExtendedRequest();
                fastTransferDestinationPutBufferExtendedRequest.RopId = (byte)RopId.RopFastTransferDestinationPutBufferExtended;
                fastTransferDestinationPutBufferExtendedRequest.LogonId = 0x00;
                fastTransferDestinationPutBufferExtendedRequest.InputHandleIndex = 0x00;
                fastTransferDestinationPutBufferExtendedRequest.TransferDataSize = (ushort)transferData.Length;
                fastTransferDestinationPutBufferExtendedRequest.TransferData = transferData;

                RopFastTransferDestinationPutBufferExtendedResponse response = (RopFastTransferDestinationPutBufferExtendedResponse)this.Process(serverId, fastTransferDestinationPutBufferExtendedRequest, fastTransferDestinationPutBufferExtendedHandle);

                if (RopResult.Success != (RopResult)response.ReturnValue)
                {
                    result = (RopResult)response.ReturnValue;
                    Site.Log.Add(LogEntryKind.Debug, "Call FastTransferDestinationPutBufferExtended fail at {0} times, return error: {1}", i, result);
                    break;
                }
            }

            return result;
        }
        /// <summary>
        ///  Initializes a FastTransfer operation for uploading content encoded in a client-provided FastTransfer stream into a mailbox
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="sourceHandleIndex">A fast transfer stream object handle index.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="copyFlag">Defines parameters of the FastTransfer download operation.</param>
        /// <param name="configHandleIndex">Configure handle's index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult FastTransferDestinationConfigure(int serverId, int sourceHandleIndex, SourceOperation option, FastTransferDestinationConfigureCopyFlags copyFlag, out int configHandleIndex)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;
            RopFastTransferDestinationConfigureRequest fastTransferDestinationConfigureRequest;
            uint fastTransferDestinationConfigureHandle = this.handleContainer[sourceHandleIndex];
            configHandleIndex = -1;

            // Construct ROP request.
            fastTransferDestinationConfigureRequest.RopId = 0x53;
            fastTransferDestinationConfigureRequest.LogonId = 0x00;
            fastTransferDestinationConfigureRequest.InputHandleIndex = 0x00;
            fastTransferDestinationConfigureRequest.OutputHandleIndex = 0x01;
            fastTransferDestinationConfigureRequest.SourceOperation = (byte)option;
            fastTransferDestinationConfigureRequest.CopyFlags = (byte)copyFlag;

            // Send request and get response.
            RopFastTransferDestinationConfigureResponse fastTransferDestinationConfigureResponse = (RopFastTransferDestinationConfigureResponse)this.Process(serverId, fastTransferDestinationConfigureRequest, fastTransferDestinationConfigureHandle);
            result = (RopResult)fastTransferDestinationConfigureResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                configHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(configHandleIndex, this.responseSOHs[fastTransferDestinationConfigureResponse.OutputHandleIndex]);

                // Verify ROP FastTransferDestinationConfigure
                this.VerifyRopFastTransferDestinationConfigure(fastTransferDestinationConfigureResponse);
            }

            return result;
        }

        /// <summary>
        /// Define the scope and parameters of the synchronization download operation. 
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderHandleIndex">The server object handle index.</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="option">Defines the parameters of a download operation.</param>
        /// <param name="synchronizationFlag">Flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="synchronizationExtraFlag">Extra flag structure that defines the parameters of the synchronization operation.</param>
        /// <param name="property">A list of properties and subobjects to exclude or include.</param>
        /// <param name="downloadcontextHandleIndex">Synchronization download context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationConfigure(int serverId, int folderHandleIndex, SynchronizationTypes synchronizationType, SendOptionAlls option, SynchronizationFlag synchronizationFlag, SynchronizationExtraFlag synchronizationExtraFlag, Sequence<string> property, out int downloadcontextHandleIndex)
        {
            // Initialize ROP data.
            downloadcontextHandleIndex = -1;
            uint objHandle = this.handleContainer[folderHandleIndex];
            RopSynchronizationConfigureRequest synchronizationConfigureRequest;
            RopResult result = RopResult.InvalidParameter;
            switch (synchronizationType)
            {
                case SynchronizationTypes.Hierarchy:
                    this.streamType = FastTransferStreamType.hierarchySync;
                    break;
                case SynchronizationTypes.Contents:
                    this.streamType = FastTransferStreamType.contentsSync;
                    break;
                default:
                    break;
            }

            PropertyTag[] propertyTags = new PropertyTag[property.Count];
            for (int i = 0; i < property.Count; i++)
            {
                propertyTags[i] = this.propertyTagsDictionary[property[i]];
            }

            this.synchroniztionFlag = synchronizationFlag;
            this.propertyTagForConfigure = propertyTags[0];

            // Construct RopSynchronizationConfigure request.
            synchronizationConfigureRequest.RopId = (byte)RopId.RopSynchronizationConfigure;
            synchronizationConfigureRequest.LogonId = 0x00;
            synchronizationConfigureRequest.InputHandleIndex = 0x00;
            synchronizationConfigureRequest.OutputHandleIndex = 0x01;

            // 0x01,Indicates a contents synchronization.0x02,Indicates a hierarchy synchronization.
            synchronizationConfigureRequest.SynchronizationType = (byte)synchronizationType;
            synchronizationConfigureRequest.SendOptions = (byte)option;

            // Unicode 
            synchronizationConfigureRequest.SynchronizationFlags = (ushort)synchronizationFlag;
            if (synchronizationType == SynchronizationTypes.Contents)
            {
                // Set RestrictionDataSize to 0x0000 to avoid the complex RestrictionData.
                synchronizationConfigureRequest.RestrictionDataSize = 0x0005;

                // If PidTagMessageClass exists 
                synchronizationConfigureRequest.RestrictionData = new byte[] { 0x08, 0x1f, 0x00, 0x1a, 0x00 };
            }
            else
            {
                // Set RestrictionDataSize to 0x0000 to avoid the complex RestrictionData.
                synchronizationConfigureRequest.RestrictionDataSize = 0x0000;
                synchronizationConfigureRequest.RestrictionData = new byte[0];
            }

            synchronizationConfigureRequest.SynchronizationExtraFlags = (byte)synchronizationExtraFlag;
            synchronizationConfigureRequest.PropertyTagCount = (ushort)propertyTags.Length;
            synchronizationConfigureRequest.PropertyTags = propertyTags;

            if (this.Process(serverId, synchronizationConfigureRequest, objHandle) is RopSynchronizationConfigureResponse)
            {
                // Send request and get response.
                RopSynchronizationConfigureResponse synchronizationConfigureResponse = (RopSynchronizationConfigureResponse)this.Process(serverId, synchronizationConfigureRequest, objHandle);
                result = (RopResult)synchronizationConfigureResponse.ReturnValue;

                if (result == RopResult.Success)
                {
                    downloadcontextHandleIndex = AdapterHelper.GetHandleIndex();
                    this.handleContainer.Add(downloadcontextHandleIndex, this.responseSOHs[synchronizationConfigureResponse.OutputHandleIndex]);

                    if ((synchronizationExtraFlag & SynchronizationExtraFlag.OrderByDeliveryTime) != SynchronizationExtraFlag.OrderByDeliveryTime)
                    {
                        this.isOrderByDeliveryTimeExtraFlagNotSet = true;
                    }

                    if (this.lastChangeMadeByServer && (synchronizationExtraFlag & SynchronizationExtraFlag.CN) == SynchronizationExtraFlag.CN)
                    {
                        this.lastChangeMadeByClient = true;
                    }
                }

                // Verify ROP SynchronizationConfigure
                this.VerifyRopSynchronizationConfigure(synchronizationConfigureRequest, synchronizationConfigureResponse);
            }
            else
            {
                // If the response type is not RopSynchronizationConfigureResponse it must be that the process method throughout a FormatException when sending RPC request buffer.
                result = RopResult.RpcFormat;
            }

            return result;
        }

        /// <summary>
        /// Upload of an ICS state property into the synchronization context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">The synchronization context handle</param>
        /// <param name="icsPropertyType">Property tags of the ICS state properties.</param>
        /// <param name="isPidTagIdsetGivenInputAsInter32"> identifies Property tags as PtypInteger32.</param>
        /// <param name="icsStateIndex">The index of the ICS State.</param>
        /// <returns>The ICS state property is upload to the server successfully or not.</returns>
        public RopResult SynchronizationUploadState(int serverId, int uploadContextHandleIndex, ICSStateProperties icsPropertyType, bool isPidTagIdsetGivenInputAsInter32, int icsStateIndex)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;
            uint synUploadContextHandle = this.handleContainer[uploadContextHandleIndex];
            ICSStateData uploadICSState = new ICSStateData();
            if (icsStateIndex != 0)
            {
                uploadICSState = this.icsStateContainer[icsStateIndex];
            }
            else
            {
                uploadICSState.PidTagCnsetRead = new byte[0];
                uploadICSState.PidTagCnsetSeen = new byte[0];
                uploadICSState.PidTagCnsetSeenFAI = new byte[0];
                uploadICSState.PidTagIdsetGiven = new byte[0];
            }

            byte[] currentICSPropertyValue;

            // Construct ROP request.
            RopSynchronizationUploadStateStreamBeginRequest synchronizationUploadStateStreamBeginRequest;
            synchronizationUploadStateStreamBeginRequest.RopId = 0x75;
            synchronizationUploadStateStreamBeginRequest.LogonId = 0x00;
            synchronizationUploadStateStreamBeginRequest.InputHandleIndex = 0x00;
            uint stateProperty = 0;

            // propertyId and propertyTypId should be joined up. So propertyId should be move left 16 positions for propertyTypeID.
            switch (icsPropertyType)
            {
                case ICSStateProperties.PidTagCnsetRead:

                    // PidTagCnsetRead propertyId 0x67DA, propertyTypeId 0x0102
                    stateProperty = (uint)((0x67D2 << 16) | 0x0102);
                    currentICSPropertyValue = new byte[uploadICSState.PidTagCnsetRead.Length];
                    currentICSPropertyValue = uploadICSState.PidTagCnsetRead;
                    break;
                case ICSStateProperties.PidTagCnsetSeen:

                    // PidTagCnsetSeen propertyId 0x6796, propertyTypeId 0x0102
                    stateProperty = (uint)((0x6796 << 16) | 0x0102);
                    currentICSPropertyValue = new byte[uploadICSState.PidTagCnsetSeen.Length];
                    currentICSPropertyValue = uploadICSState.PidTagCnsetSeen;
                    break;
                case ICSStateProperties.PidTagCnsetSeenFAI:

                    // PidTagCnsetSeenFAI propertyId 0x67DA, propertyTypeId 0x0102
                    stateProperty = (uint)((0x67DA << 16) | 0x0102);
                    currentICSPropertyValue = new byte[uploadICSState.PidTagCnsetSeenFAI.Length];
                    currentICSPropertyValue = uploadICSState.PidTagCnsetSeenFAI;
                    break;
                case ICSStateProperties.PidTagIdsetGiven:

                    // PidTagIdsetGiven propertyId 0x4017, propertyTypeId 0x003
                    stateProperty = (uint)((0x4017 << 16) | 0x0102);
                    if (isPidTagIdsetGivenInputAsInter32)
                    {
                        stateProperty = (uint)((0x4017 << 16) | 0x0003);
                    }

                    currentICSPropertyValue = new byte[uploadICSState.PidTagIdsetGiven.Length];
                    currentICSPropertyValue = uploadICSState.PidTagIdsetGiven;
                    break;
                default:
                    currentICSPropertyValue = new byte[0];
                    break;
            }

            synchronizationUploadStateStreamBeginRequest.TransferBufferSize = (uint)currentICSPropertyValue.Length;
            synchronizationUploadStateStreamBeginRequest.StateProperty = stateProperty;

            // Send the RopSynchronizationUploadStateStreamBegin request and get response from server.
            RopSynchronizationUploadStateStreamBeginResponse synchronizationUploadStateStreamBeginResponse = (RopSynchronizationUploadStateStreamBeginResponse)this.Process(serverId, synchronizationUploadStateStreamBeginRequest, synUploadContextHandle);
            result = (RopResult)synchronizationUploadStateStreamBeginResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SynchronizationUploadStateStreamBegin
                this.VerifyRopSynchronizationUploadStateStreamBegin(synchronizationUploadStateStreamBeginResponse);
                if (isPidTagIdsetGivenInputAsInter32)
                {
                    this.VerifyServerAcceptPidTagIdsetGivenPtypInteger32(result);
                }
            }
            else
            {
                return result;
            }

            // Construct the RopSynchronizationUploadStateStreamContinue request.
            RopSynchronizationUploadStateStreamContinueRequest synchronizationUploadStateStreamContinueRequest;
            synchronizationUploadStateStreamContinueRequest.RopId = 0x76;
            synchronizationUploadStateStreamContinueRequest.LogonId = 0x00;
            synchronizationUploadStateStreamContinueRequest.InputHandleIndex = 0x00;
            synchronizationUploadStateStreamContinueRequest.StreamDataSize = (uint)currentICSPropertyValue.Length;
            synchronizationUploadStateStreamContinueRequest.StreamData = currentICSPropertyValue;

            // Send the RopSynchronizationUploadStateStreamContinue request.
            RopSynchronizationUploadStateStreamContinueResponse synchronizationUploadStateStreamContinueResponse = (RopSynchronizationUploadStateStreamContinueResponse)this.Process(serverId, synchronizationUploadStateStreamContinueRequest, synUploadContextHandle);
            result = (RopResult)synchronizationUploadStateStreamContinueResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SynchronizationUploadStateStreamContinue
                this.VerifyRopSynchronizationUploadStateStreamContinue(synchronizationUploadStateStreamContinueResponse);
            }
            else
            {
                return result;
            }

            // Construct the RopSynchronizationUploadStateStreamEnd request.
            RopSynchronizationUploadStateStreamEndRequest synchronizationUploadStateStreamEndRequest;
            synchronizationUploadStateStreamEndRequest.RopId = 0x77;
            synchronizationUploadStateStreamEndRequest.LogonId = 0x00;
            synchronizationUploadStateStreamEndRequest.InputHandleIndex = 0x00;

            RopSynchronizationUploadStateStreamEndResponse synchronizationUploadStateStreamEndResponse = (RopSynchronizationUploadStateStreamEndResponse)this.Process(serverId, synchronizationUploadStateStreamEndRequest, synUploadContextHandle);
            result = (RopResult)synchronizationUploadStateStreamEndResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SynchronizationUploadStateStreamEnd
                this.VerifyRopSynchronizationUploadStateStreamEnd(synchronizationUploadStateStreamEndResponse);
            }

            return result;
        }

        /// <summary>
        /// Configures the synchronization upload operation.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index .</param>
        /// <param name="synchronizationType">The type of synchronization requested: contents or hierarchy.</param>
        /// <param name="synchronizationHandleIndex">Synchronization upload context handle index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationOpenCollector(int serverId, int objHandleIndex, SynchronizationTypes synchronizationType, out int synchronizationHandleIndex)
        {
            // Initialize ROP data.
            synchronizationHandleIndex = -1;
            uint targetFolderHandle = this.handleContainer[objHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopSynchronizationOpenCollectorRequest synchronizationOpenCollectorRequest;
            synchronizationOpenCollectorRequest.RopId = (byte)RopId.RopSynchronizationOpenCollector;
            synchronizationOpenCollectorRequest.LogonId = 0x00;
            synchronizationOpenCollectorRequest.InputHandleIndex = 0x00;
            synchronizationOpenCollectorRequest.OutputHandleIndex = 0x01;
            if (synchronizationType == SynchronizationTypes.Contents)
            {
                // Synchronization upload is requested for contents of folders
                synchronizationOpenCollectorRequest.IsContentsCollector = 0x01;
            }
            else
            {
                synchronizationOpenCollectorRequest.IsContentsCollector = 0x00;
            }

            // Send request and get response.
            RopSynchronizationOpenCollectorResponse synchronizationOpenCollectorResponse = (RopSynchronizationOpenCollectorResponse)this.Process(serverId, synchronizationOpenCollectorRequest, targetFolderHandle);
            result = (RopResult)synchronizationOpenCollectorResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                synchronizationHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(synchronizationHandleIndex, this.responseSOHs[synchronizationOpenCollectorResponse.OutputHandleIndex]);

                // Verify ROP SynchronizationOpenCollector
                this.VerifyRopSynchronizationOpenCollector(synchronizationOpenCollectorResponse);
            }

            return result;
        }

        /// <summary>
        /// Import new folders, or changes to existing folders, into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Upload context handle.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="properties">Properties to be set.</param>
        /// <param name="localFolderIdIndex">Local folder id index</param>
        /// <param name="folderIdIndex">The folder object id index.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportHierarchyChange(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex)
        {
            // Initialize ROP data.
            uint uploadContextHandle = this.handleContainer[uploadContextHandleIndex];
            RopResult result = RopResult.InvalidParameter;
            folderIdIndex = -1;
            ulong folderId;

            // Construct ROP request.
            RopSynchronizationImportHierarchyChangeRequest synchronizationImportHierarchyChangeRequest;

            // Construct the RopSynchronizationImportHierarchyChange request.
            synchronizationImportHierarchyChangeRequest.RopId = 0x73;
            synchronizationImportHierarchyChangeRequest.LogonId = 0x00;
            synchronizationImportHierarchyChangeRequest.InputHandleIndex = 0x00;

            // Get the folder property value.
            TaggedPropertyValue[] taggedPropertyValueArrays = new TaggedPropertyValue[properties.Count];
            int index = 0;
            foreach (string tempProperty in properties)
            {
                taggedPropertyValueArrays[index++] = this.taggedPropertyValuesDictionary[tempProperty];
            }

            int tempParentFolderHandleIndex = parentFolderHandleIndex == -1 ? 3 : parentFolderHandleIndex;
            TaggedPropertyValue[] hierarchyValues = this.CreateSampleHierarchyValues(serverId, localFolderIdIndex, tempParentFolderHandleIndex, out folderId, ConflictTypes.NONE);
            if (parentFolderHandleIndex == -1)
            {
                hierarchyValues[0].Value[12] -= 1;
            }

            synchronizationImportHierarchyChangeRequest.HierarchyValueCount = (ushort)hierarchyValues.Length;
            synchronizationImportHierarchyChangeRequest.HierarchyValues = hierarchyValues;
            synchronizationImportHierarchyChangeRequest.PropertyValueCount = (ushort)taggedPropertyValueArrays.Length;
            synchronizationImportHierarchyChangeRequest.PropertyValues = taggedPropertyValueArrays;

            // Send request and get response.
            RopSynchronizationImportHierarchyChangeResponse synchronizationImportHierarchyChangeResponse = (RopSynchronizationImportHierarchyChangeResponse)this.Process(serverId, synchronizationImportHierarchyChangeRequest, uploadContextHandle);
            result = (RopResult)synchronizationImportHierarchyChangeResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                if (!this.objectIdContainer.ContainsKey(localFolderIdIndex))
                {
                    folderIdIndex = AdapterHelper.GetObjectIdIndex();
                    this.objectIdContainer.Add(folderIdIndex, folderId);
                }

                // Verify ROP SynchronizationImportHierarchyChange
                this.VerifyRopSynchronizationImportHierarchyChange(synchronizationImportHierarchyChangeResponse);
            }

            return result;
        }

        /// <summary>
        /// Import new messages or changes to existing messages into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">A synchronization upload context handle index.</param>
        /// <param name="localMessageidIndex">Message Id index.</param>
        /// <param name="importFlag">An 8-bit flag .</param>
        /// <param name="importMessageHandleIndex">The index of handle that indicate the Message object into which the client will upload the rest of the message changes.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportMessageChange(int serverId, int uploadContextHandleIndex, int localMessageidIndex, ImportFlag importFlag, out int importMessageHandleIndex)
        {
            // Initialize ROP data.
            importMessageHandleIndex = -1;
            uint synchronizationImportMessageChangeHandle = this.handleContainer[uploadContextHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopSynchronizationImportMessageChangeRequest synchronizationImportMessageChangeRequest;

            // Synchronization ImportHierarchy Change.
            TaggedPropertyValue[] messagePropertyValues = this.CreateSamplePropertyValues(localMessageidIndex);

            // Construct the RopSynchronizationImportMessageChange request.
            synchronizationImportMessageChangeRequest.RopId = 0x72;
            synchronizationImportMessageChangeRequest.LogonId = 0x00;
            synchronizationImportMessageChangeRequest.InputHandleIndex = 0x00;
            synchronizationImportMessageChangeRequest.OutputHandleIndex = 0x01;
            synchronizationImportMessageChangeRequest.ImportFlag = (byte)importFlag;
            synchronizationImportMessageChangeRequest.PropertyValueCount = (ushort)messagePropertyValues.Length;
            synchronizationImportMessageChangeRequest.PropertyValues = messagePropertyValues;

            // Send ROP request and get response.
            RopSynchronizationImportMessageChangeResponse synchronizationImportMessageChangeResponse = (RopSynchronizationImportMessageChangeResponse)this.Process(serverId, synchronizationImportMessageChangeRequest, synchronizationImportMessageChangeHandle);
            result = (RopResult)synchronizationImportMessageChangeResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                importMessageHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(importMessageHandleIndex, this.responseSOHs[synchronizationImportMessageChangeResponse.OutputHandleIndex]);
                this.importFlag = importFlag;

                // Verify ROP SynchronizationImportMessageChange
                this.VerifyRopSynchronizationImportMessageChange(synchronizationImportMessageChangeResponse);
                this.importMessageTimes += 1;
            }

            return result;
        }

        /// <summary>
        /// Creates a FastTransfer download context for a snapshot of the checkpoint ICS state of the operation identified by the given synchronization download context, or synchronization upload context.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Synchronization context index.</param>
        /// <param name="stateHandleIndex">The index of FastTransfer download context for the ICS state.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationGetTransferState(int serverId, int objHandleIndex, out int stateHandleIndex)
        {
            // Initialize ROP data.
            stateHandleIndex = -1;
            uint synchronizationGetTransferStateHandle = this.handleContainer[objHandleIndex];
            RopResult result = RopResult.InvalidParameter;
            this.streamType = FastTransferStreamType.state;

            // Construct ROP request.
            RopSynchronizationGetTransferStateRequest synchronizationGetTransferStateRequest;
            synchronizationGetTransferStateRequest.RopId = 0x82;
            synchronizationGetTransferStateRequest.LogonId = 0x00;
            synchronizationGetTransferStateRequest.InputHandleIndex = 0x00;
            synchronizationGetTransferStateRequest.OutputHandleIndex = 0x01;

            // Send the CreateSyncUploadContext request and verify the success response.
            RopSynchronizationGetTransferStateResponse synchronizationGetTransferStateResponse = (RopSynchronizationGetTransferStateResponse)this.Process(serverId, synchronizationGetTransferStateRequest, synchronizationGetTransferStateHandle);
            result = (RopResult)synchronizationGetTransferStateResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                stateHandleIndex = AdapterHelper.GetHandleIndex();
                this.handleContainer.Add(stateHandleIndex, this.responseSOHs[synchronizationGetTransferStateResponse.OutputHandleIndex]);

                // Verify ROP SynchronizationGetTransferState
                this.VerifyRopSynchronizationGetTransferState(synchronizationGetTransferStateResponse);
            }

            return result;
        }

        /// <summary>
        /// Imports deletions of messages or folders into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server</param>
        /// <param name="uploadcontextHandleIndex">Synchronization upload context handle</param>
        /// <param name="objIdIndexes">more object id</param>
        /// <param name="importDeleteFlag">Deletions type</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportDeletes(int serverId, int uploadcontextHandleIndex, Sequence<int> objIdIndexes, byte importDeleteFlag)
        {
            // Initialize ROP data.
            uint synchronizationImportDeletesHandle = this.handleContainer[uploadcontextHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            RopSynchronizationImportDeletesRequest synchronizationImportDeletesRequest = new RopSynchronizationImportDeletesRequest();
            RopSynchronizationImportDeletesResponse synchronizationImportDeletesResponse = new RopSynchronizationImportDeletesResponse();

            // Construct the RopSynchronizationImportDeletes request.
            synchronizationImportDeletesRequest.RopId = 0x74;
            synchronizationImportDeletesRequest.InputHandleIndex = 0x00;
            synchronizationImportDeletesRequest.LogonId = 0x00;
            synchronizationImportDeletesRequest.IsHierarchy = (byte)importDeleteFlag;
            this.importDelFlag = importDeleteFlag;
            synchronizationImportDeletesRequest.PropertyValueCount = (ushort)objIdIndexes.Count;
            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[objIdIndexes.Count];
            for (int i = 0; i < propertyValues.Length; i++)
            {
                TaggedPropertyValue propertyValue = new TaggedPropertyValue();
                PropertyTag[] tagArray = new PropertyTag[1];
                byte[] deleteGidValue = new byte[sizeof(int) + sizeof(short) + GidLength];
                propertyValue.PropertyTag = new PropertyTag(0x0000, 0x1102);

                // The combination of first two bytes (0x0016) indicates the length of value field.
                // The amount of id
                int index = 0;
                Array.Copy(BitConverter.GetBytes(tagArray.Length), 0, deleteGidValue, 0, sizeof(int));
                index += sizeof(int);
                Array.Copy(BitConverter.GetBytes((short)GidLength), 0, deleteGidValue, index, sizeof(short));
                index += sizeof(short);
                ulong gid = this.objectIdContainer[objIdIndexes[i]];
                this.delObjId = gid;
                byte[] longTermId = this.GetLongTermIdByte(serverId, gid);
                Array.Copy(longTermId, 0, deleteGidValue, index, longTermId.Length);
                propertyValue.Value = deleteGidValue;
                propertyValues[i] = propertyValue;
            }

            synchronizationImportDeletesRequest.PropertyValues = propertyValues;

            // Send the RopSynchronizationImportDeletes request to get the success response.
            synchronizationImportDeletesResponse = (RopSynchronizationImportDeletesResponse)this.Process(serverId, synchronizationImportDeletesRequest, synchronizationImportDeletesHandle);
            result = (RopResult)synchronizationImportDeletesResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                this.isNonImportMessageChangeOperation = true;
                this.hasExecuteSynchronizationImportDeletes = true;
            }

            // Verify ROP SynchronizationImportDeletes
            this.VerifyRopSynchronizationImportDeletes(synchronizationImportDeletesRequest, synchronizationImportDeletesResponse);

            return result;
        }

        /// <summary>
        /// Allocates a range of internal identifiers for the purpose of assigning them to client-originated objects in a local replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">The server object handle index.</param>
        /// <param name="idcount">An unsigned 32-bit integer specifies the number of IDs to allocate.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult GetLocalReplicaIds(int serverId, int objHandleIndex, uint idcount)
        {
            // Initialize ROP data.
            uint getLocalReplicaIdsHandle = this.handleContainer[objHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopGetLocalReplicaIdsRequest ropGetLocalReplicaIdsRequest;
            ropGetLocalReplicaIdsRequest.IdCount = idcount;
            ropGetLocalReplicaIdsRequest.InputHandleIndex = 0x00;
            ropGetLocalReplicaIdsRequest.LogonId = 0x00;
            ropGetLocalReplicaIdsRequest.RopId = 0x7F;

            // Send request and get response.
            RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse = (RopGetLocalReplicaIdsResponse)this.Process(serverId, ropGetLocalReplicaIdsRequest, getLocalReplicaIdsHandle);
            result = (RopResult)ropGetLocalReplicaIdsResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                this.serverReplicaGuid = new Guid(ropGetLocalReplicaIdsResponse.ReplGuid);
                this.localId = ropGetLocalReplicaIdsResponse.GlobalCount;

                // Verify ROP GetLocalReplicaIds
                this.VerifyRopGetLocalReplicaIds(ropGetLocalReplicaIdsResponse);
            }

            return result;
        }

        /// <summary>
        /// Imports message read state changes into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="objHandleIndex">Sync handle.</param>
        /// <param name="objectHandleIndex">Message handle or folder handle or attachments handle.</param>
        /// <param name="readStatus">A boolean value indicating the message read status, true means read.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportReadStateChanges(int serverId, int objHandleIndex, int objectHandleIndex, bool readStatus)
        {
            uint synchronizationImportReadStateChangesHandle = this.handleContainer[objHandleIndex];
            uint messageHandle = this.handleContainer[objectHandleIndex];
            RopResult result = RopResult.InvalidParameter;

            // Construct RopGetPropertiesSpecificRequest.
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                PropertySizeLimit = 0xFFFF
            };

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned,
            PropertyTag[] tagArray = new PropertyTag[1];
            PropertyTag tag = new PropertyTag(0x65E0, 0x0102);
            tagArray[0] = tag;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = tagArray;

            // Get specific properties value.
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.Process(serverId, getPropertiesSpecificRequest, messageHandle);

            MessageReadState[] messageReadStates = new MessageReadState[1];
            MessageReadState messageReadState = new MessageReadState
            {
                MarkAsRead = readStatus ? (byte)0x01 : (byte)0x00,
                MessageId = new byte[22]
            };
            Array.Copy(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 2, messageReadState.MessageId, 0, 22);
            messageReadState.MessageIdSize = (ushort)(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value.Length - 2);
            messageReadStates[0] = messageReadState;

            // Construct the RopSynchronizationImportReadStateChanges request.
            RopSynchronizationImportReadStateChangesRequest synchronizationImportReadStateChangesRequest = new RopSynchronizationImportReadStateChangesRequest
            {
                RopId = 0x80,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                MessageReadStates = messageReadStates,
                MessageReadStateSize = (ushort)messageReadStates[0].Size()
            };

            RopSynchronizationImportReadStateChangesResponse synchronizationImportReadStateChangesResponse = (RopSynchronizationImportReadStateChangesResponse)this.Process(serverId, synchronizationImportReadStateChangesRequest, synchronizationImportReadStateChangesHandle);
            result = (RopResult)synchronizationImportReadStateChangesResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SynchronizationImportReadStateChanges
                this.VerifyRopSynchronizationImportReadStateChanges(synchronizationImportReadStateChangesResponse);
            }

            return result;
        }

        /// <summary>
        /// Imports information about moving a message between two existing folders within the same mailbox.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="synchronizationUploadContextHandleIndex">The index of the synchronization upload context configured for collecting changes to the contents of the message move destination folder.</param>
        /// <param name="sourceFolderIdIndex">The index of the source folder id in object id container.</param>
        /// <param name="destinationFolderIdIndex">The index of the destination folder id in object id container.</param>
        /// <param name="sourceMessageIdIndex">The index of source message id in object id container.</param>
        /// <param name="sourceFolderHandleIndex">The index of source folder handle in handleContainer.</param>
        /// <param name="destinationFolderHandleIndex">The index of destination folder handle in handle container.</param>
        /// <param name="inewerClientChange">If the client has a newer message.</param>
        /// <param name="iolderversion">If the server have an older version of a message .</param>
        /// <param name="icnpc">Verify if the change number has been used.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportMessageMove(int serverId, int synchronizationUploadContextHandleIndex, int sourceFolderIdIndex, int destinationFolderIdIndex, int sourceMessageIdIndex, int sourceFolderHandleIndex, int destinationFolderHandleIndex, bool inewerClientChange, out bool iolderversion, out bool icnpc)
        {
            // Initialize return value.
            iolderversion = false;
            icnpc = true;
            RopResult result = RopResult.InvalidParameter;
            uint synchronizationUploadContextHandle = this.handleContainer[synchronizationUploadContextHandleIndex];
            uint sourceFolderHandle = this.handleContainer[sourceFolderHandleIndex];
            ulong sourceFID = this.objectIdContainer[sourceFolderIdIndex];
            ulong sourceMID = this.objectIdContainer[sourceMessageIdIndex];

            // Construct the RopSynchronizationImportMessageMove request. 
            RopSynchronizationImportMessageMoveRequest importMessageMoveRequest = new RopSynchronizationImportMessageMoveRequest
            {
                RopId = 0x78,
                LogonId = 0x00,
                InputHandleIndex = 0x00
            };

            // Get source folder id.
            byte[] fidLongTermId = this.GetLongTermIdByte(serverId, sourceFID);
            importMessageMoveRequest.SourceFolderIdSize = (uint)fidLongTermId.Length;
            byte[] sourceFolderId = new byte[fidLongTermId.Length];
            Array.Copy(fidLongTermId, 0, sourceFolderId, 0, fidLongTermId.Length);
            importMessageMoveRequest.SourceFolderId = sourceFolderId;

            // Get source message id
            byte[] midLongTermId = this.GetLongTermIdByte(serverId, sourceMID);
            importMessageMoveRequest.SourceMessageIdSize = (uint)midLongTermId.Length;
            byte[] sourceMessageId = new byte[midLongTermId.Length];
            Array.Copy(midLongTermId, 0, sourceMessageId, 0, midLongTermId.Length);
            importMessageMoveRequest.SourceMessageId = sourceMessageId;

            // Get destination message id
            byte[] sub = new byte[4];
            byte[] currentLocalId = this.localId;

            // Get local id.
            Array.Copy(currentLocalId, 2, sub, 0, currentLocalId.Length - 2);
            Array.Reverse(sub);
            int tempLocalId = BitConverter.ToInt32(sub, 0);
            tempLocalId += this.localIdOffSet;
            sub = BitConverter.GetBytes(tempLocalId);
            Array.Reverse(sub);
            Array.Copy(sub, 0, currentLocalId, 2, currentLocalId.Length - 2);
            this.localIdOffSet++;
            importMessageMoveRequest.DestinationMessageIdSize = (uint)(this.serverReplicaGuid.ToByteArray().Length + currentLocalId.Length);
            byte[] destinationMessageId = new byte[this.serverReplicaGuid.ToByteArray().Length + currentLocalId.Length];
            Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, destinationMessageId, 0, this.serverReplicaGuid.ToByteArray().Length);
            Array.Copy(currentLocalId, 0, destinationMessageId, this.serverReplicaGuid.ToByteArray().Length, currentLocalId.Length);
            LongTermId longtermid = new LongTermId
            {
                DatabaseGuid = this.serverReplicaGuid.ToByteArray(),
                GlobalCounter = currentLocalId
            };

            // Get Mid from long termid.
            ulong destinationMID = this.GetIdFromLongTerm(serverId, longtermid);
            importMessageMoveRequest.DestinationMessageId = destinationMessageId;

            // Change number
            byte[] tempChangeNumber = this.GetChangeNumber(serverId, sourceMID, sourceFID, sourceFolderHandle);
            importMessageMoveRequest.ChangeNumberSize = (uint)tempChangeNumber.Length;

            // Generate a new change number.
            byte[] tempBytes = new byte[4];
            Array.Copy(tempChangeNumber, tempChangeNumber.Length - 4, tempBytes, 0, 4);
            Array.Reverse(tempBytes);
            int num = BitConverter.ToInt32(tempBytes, 0);
            if (inewerClientChange == false)
            {
                // Make the message version older in client than that in server.
                num--;
                iolderversion = false;
            }
            else
            {
                // Make the message version newer in client than that in server.
                num++;
                iolderversion = true;
            }

            tempBytes = BitConverter.GetBytes(num);
            Array.Reverse(tempBytes);
            Array.Copy(tempBytes, 0, tempChangeNumber, tempChangeNumber.Length - 4, tempBytes.Length);
            importMessageMoveRequest.ChangeNumber = tempChangeNumber;

            // Set the PCL size value.
            importMessageMoveRequest.PredecessorChangeListSize = (uint)(tempChangeNumber.Length + 1);
            importMessageMoveRequest.PredecessorChangeList = new byte[tempChangeNumber.Length + 1];
            importMessageMoveRequest.PredecessorChangeList[0] = (byte)tempChangeNumber.Length;
            Array.Copy(tempChangeNumber, 0, importMessageMoveRequest.PredecessorChangeList, 1, tempChangeNumber.Length);

            // Send request and get response.
            RopSynchronizationImportMessageMoveResponse synchronizationImportMessageMoveResponse = (RopSynchronizationImportMessageMoveResponse)this.Process(serverId, importMessageMoveRequest, synchronizationUploadContextHandle);
            result = (RopResult)synchronizationImportMessageMoveResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SynchronizationImportMessageMove
                this.VerifyRopSynchronizationImportMessageMove(synchronizationImportMessageMoveResponse);
                int messageIdIndex = AdapterHelper.GetObjectIdIndex();
                this.objectIdContainer.Add(messageIdIndex, destinationMID);
                icnpc = false;
                this.isNonImportMessageChangeOperation = true;
            }
            else if (result == RopResult.NewerClientChange)
            {
                // Verify ROP SynchronizationImportMessageMove
                this.VerifyRopSynchronizationImportMessageMove(synchronizationImportMessageMoveResponse);
            }

            this.isImportMessageMoveROP = true;

            return result;
        }

        /// <summary>
        /// Identifies that a set of IDs either belongs to deleted messages in the specified folder or will never be used for any messages in the specified folder.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="folderhandleIndex">A Folder object handle index.</param>
        /// <param name="longTermIdRangeIndex">An array of LongTermIdRange structures defines a range of IDs, which are reported as unused or deleted.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SetLocalReplicaMidsetDeleted(int serverId, int folderhandleIndex, Sequence<int> longTermIdRangeIndex)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;
            uint folderHandle = this.handleContainer[folderhandleIndex];
            LongTermIdRange[] idranges = new LongTermIdRange[longTermIdRangeIndex.Count];

            // Get the Mid need to be deleted.
            for (int i = 0; i < longTermIdRangeIndex.Count; i++)
            {
                LongTermId max = new LongTermId();
                LongTermId min = new LongTermId();

                // Max Mid
                max.DatabaseGuid = this.serverReplicaGuid.ToByteArray();
                max.GlobalCounter = new byte[6];
                Array.Copy(this.localId, 0, max.GlobalCounter, 0, 6);
                byte[] sub = new byte[4];
                Array.Copy(max.GlobalCounter, 2, sub, 0, 4);
                Array.Reverse(sub);
                int num = BitConverter.ToInt32(sub, 0);

                // Make the Mid to the mid to be deleted one.
                num += longTermIdRangeIndex[i];
                sub = BitConverter.GetBytes(num);
                Array.Reverse(sub);
                Array.Copy(sub, 0, max.GlobalCounter, 2, 4);

                // Min Mid
                min.DatabaseGuid = this.serverReplicaGuid.ToByteArray();
                min.GlobalCounter = new byte[6];
                Array.Copy(this.localId, 0, min.GlobalCounter, 0, 6);
                sub = new byte[4];
                Array.Copy(min.GlobalCounter, 2, sub, 0, 4);
                Array.Reverse(sub);
                num = BitConverter.ToInt32(sub, 0);

                // Make the Mid to the mid to be deleted one.
                num += longTermIdRangeIndex[i];
                sub = BitConverter.GetBytes(num);
                Array.Reverse(sub);
                Array.Copy(sub, 0, min.GlobalCounter, 2, 4);

                // ID range data
                idranges[i].MaxLongTermId = max.Serialize();
                idranges[i].MinLongTermId = min.Serialize();
            }

            // Construct the RopGetLocalReplicaIds request.
            RopSetLocalReplicaMidsetDeletedRequest req = new RopSetLocalReplicaMidsetDeletedRequest
            {
                RopId = 0x93
            };
            byte logonId = 0;
            req.LogonId = logonId;
            req.InputHandleIndex = 0;
            req.DataSize = (ushort)((sizeof(byte) * 48 * idranges.Length) + 4);
            req.LongTermIdRangeCount = (uint)idranges.Length;

            // Set idCount, which specifies the number of IDs to reserve,
            req.LongTermIdRanges = idranges;

            // Send request and get response.
            RopSetLocalReplicaMidsetDeletedResponse setLocalReplicaMidsetDeletedResponse = (RopSetLocalReplicaMidsetDeletedResponse)this.Process(serverId, req, folderHandle);
            result = (RopResult)setLocalReplicaMidsetDeletedResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                // Verify ROP SetLocalReplicaMidsetDeleted
                this.VerifyRopSetLocalReplicaMidsetDeleted(setLocalReplicaMidsetDeletedResponse);
            }

            return result;
        }

        /// <summary>
        /// Get specific property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleindex">Identify from which the property will be gotten.</param>
        /// <param name="propertyTag">A list of propertyTags.</param>
        /// <returns>Indicate the result of this ROP operation. </returns>
        public RopResult GetPropertiesSpecific(int serverId, int handleindex, Sequence<string> propertyTag)
        {
            // Initialize ROP data.
            uint handle = this.handleContainer[handleindex];
            RopResult result = RopResult.InvalidParameter;
            PropertyTag[] propertyTags = new PropertyTag[propertyTag.Count];

            // Initial propertyTag.
            for (int i = 0; i < propertyTag.Count; i++)
            {
                propertyTags[i] = this.propertyTagsDictionary[propertyTag[i]];
            }

            // Generate ROP request.
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                PropertySizeLimit = 0xFFFF
            };

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned,
            PropertyTag[] tagArray = propertyTags;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = tagArray;

            // Send the RopGetPropertiesSpecific request
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.Process(serverId, getPropertiesSpecificRequest, handle);
            result = (RopResult)getPropertiesSpecificResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                for (int i = 0; i < propertyTag.Count; i++)
                {
                    if (this.propertyValuesSpecific.ContainsKey(propertyTag[i]))
                    {
                        // If the property has been already existed modify it, otherwise add it.
                        this.propertyValuesSpecific[propertyTag[i]] = getPropertiesSpecificResponse.RowData.PropertyValues[i].Serialize();
                    }
                    else
                    {
                        this.propertyValuesSpecific.Add(propertyTag[i], getPropertiesSpecificResponse.RowData.PropertyValues[i].Serialize());
                    }

                    if (propertyTag[i] == "PidTagResolveMethod")
                    {
                        this.VerifyPidTagResolveMethodValue();
                    }
                    else if (propertyTag[i] == "PidTagMessageFlags")
                    {
                        this.VerifySyncImportReadStateChanged(getPropertiesSpecificResponse.RowData.PropertyValues[i]);
                    }
                    else if (propertyTag[i] == "PidTagChangeNumber" || propertyTag[i] == "PidTagDeletedCountTotal" || propertyTag[i] == "PidTagLocalCommitTimeMax")
                    {
                        this.VerifyPidTagChangeNumberValueRelated(propertyTag[i], getPropertiesSpecificResponse.RowData.PropertyValues[i]);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Set the specific object's property value.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="handleindex">Server object handle index.</param>
        /// <param name="taggedPropertyValueArray">The list of property values.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SetProperties(int serverId, int handleindex, Sequence<string> taggedPropertyValueArray)
        {
            // Initialize ROP data.
            TaggedPropertyValue[] taggedPropertyValueArrays = new TaggedPropertyValue[taggedPropertyValueArray.Count];

            // Set a large property to test TransferStatus (Partial and NoRoom)
            for (int i = 0; i < taggedPropertyValueArray.Count; i++)
            {
                if (taggedPropertyValueArray[i] == "PidTagPolicyTag")
                {
                    // Set the property PidTagPolicyTag
                    taggedPropertyValueArrays[i] = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x3019, 0x0102)
                    };
                    string value = string.Empty;
                    for (int j = 0; j < 30000; j++)
                    {
                        value += "a";
                    }

                    taggedPropertyValueArrays[i].Value = Common.AddInt16LengthBeforeBinaryArray(Encoding.ASCII.GetBytes(value));
                }
                else
                {
                    taggedPropertyValueArrays[i] = this.taggedPropertyValuesDictionary[taggedPropertyValueArray[i]];
                }
            }

            uint setPropertiesHandle = this.handleContainer[handleindex];
            RopResult result = RopResult.InvalidParameter;

            // Construct ROP request.
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest
            {
                RopId = 0x0A,
                LogonId = 0x00,
                InputHandleIndex = 0x00
            };
            int arraySize = 0;
            for (int i = 0; i < taggedPropertyValueArrays.Length; i++)
            {
                arraySize += taggedPropertyValueArrays[i].Size();
            }

            // This value specifies the number of bytes used for the PropertyValueCount field and the PropertyValues field.
            setPropertiesRequest.PropertyValueSize = (ushort)(arraySize + 2);
            setPropertiesRequest.PropertyValueCount = (ushort)taggedPropertyValueArrays.Length;
            setPropertiesRequest.PropertyValues = taggedPropertyValueArrays;

            // Send the RopSetProperties request 
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.Process(serverId, setPropertiesRequest, setPropertiesHandle);
            result = (RopResult)setPropertiesResponse.ReturnValue;
            if (result == RopResult.Success)
            {
                this.lastChangeMadeByServer = true;
            }

            return result;
        }

        /// <summary>
        /// Modifies the permissions associated with a folder.
        /// </summary>
        /// <param name="serverId">The server id</param>
        /// <param name="folderHandleIndex">index of folder handle in container</param>
        /// <param name="permissionLevel">The permission level</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult ModifyPermissions(int serverId, int folderHandleIndex, PermissionLevels permissionLevel)
        {
            // Initialize ROP data.
            RopResult result = RopResult.InvalidParameter;
            uint folderHandle = this.handleContainer[folderHandleIndex];

            // Add Administrator user into permission list of the specific folder.
            TaggedPropertyValue[] taggedProperties = new TaggedPropertyValue[2];

            // EntryId
            TaggedPropertyValue propertyValueEntryId = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x0FFF, 0x0102)
            };
            string userDN = Common.GetConfigurationPropertyValue("AdminUserESSDN", this.Site);
            AddressBookEntryId addressEntryId = new AddressBookEntryId(userDN);
            propertyValueEntryId.Value = Common.AddInt16LengthBeforeBinaryArray(addressEntryId.Serialize());
            taggedProperties[0] = propertyValueEntryId;

            // PidTagMemberRights
            TaggedPropertyValue propertyValueMemberRight = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x6673, 0x0003),
                Value = BitConverter.GetBytes((uint)permissionLevel)
            };

            // Set permission.
            taggedProperties[1] = propertyValueMemberRight;
            PermissionData[] permissionsDataArray = new PermissionData[1];

            // Add row
            permissionsDataArray[0].PermissionDataFlags = (byte)0x01;
            permissionsDataArray[0].PropertyValueCount = (ushort)taggedProperties.Length;
            permissionsDataArray[0].PropertyValues = taggedProperties;

            // Construct ROP request.
            RopModifyPermissionsRequest modifyPermission = new RopModifyPermissionsRequest
            {
                RopId = 0x40,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                ModifyFlags = (byte)ModifyFlags.IncludeFreeBusy,
                ModifyCount = 0x01,
                PermissionsData = permissionsDataArray
            };

            // Send request and get response.
            RopModifyPermissionsResponse modifyPermissionresponse = (RopModifyPermissionsResponse)this.Process(serverId, modifyPermission, folderHandle);
            result = (RopResult)modifyPermissionresponse.ReturnValue;
            if (result == RopResult.Success)
            {
                this.currentPermission = permissionLevel;
                if (permissionLevel == PermissionLevels.None)
                {
                    // If the specific folder's permission is None, set existNoPermissionFolder as true to make client has no permission to access this folder now
                    this.existNoPermissionFolder = true;
                }
            }

            return result;
        }

        #endregion IMS_OXCFXICSAdapter Members

        /// <summary>
        /// Validate if the given two buffers are equal
        /// </summary>
        /// <param name="operation">Fast transfer operation</param>
        /// <param name="firstBufferIndex">The first buffer's index</param>
        /// <param name="secondBufferIndex">The second buffer's index</param>
        /// <returns>Returns true only if the two buffers are equal</returns>
        public bool AreEqual(EnumFastTransferOperation operation, int firstBufferIndex, int secondBufferIndex)
        {
            if (firstBufferIndex <= 0 || secondBufferIndex <= 0)
            {
                return false;
            }
            else
            {
                byte[] firstBuffer = this.streamBufferContainer[firstBufferIndex];
                byte[] secondBuffer = this.streamBufferContainer[secondBufferIndex];
                if (firstBuffer.Length != secondBuffer.Length)
                {
                    return false;
                }
                else
                {
                    for (int i = 0; i < firstBuffer.Length; i++)
                    {
                        if (firstBuffer[i] != secondBuffer[i])
                        {
                            return false;
                        }
                    }

                    return true;
                }
            }
        }

        /// <summary>
        /// Determines if the requirement is enabled or not.
        /// </summary>
        /// <param name="rsid">Requirement id.</param>
        /// <param name="enabled">Requirement is enable or not.</param>
        public void CheckRequirementEnabled(int rsid, out bool enabled)
        {
            // Add logic here        
            enabled = Common.IsRequirementEnabled(rsid, this.Site);

            // If the requirement number is 1111 with false value, which means the server doesn't support public folder replica information and the test case actually doesn't verify 
            // any requirements, so the case will report the inconclusive status.
            if (rsid == 1111 && enabled == false)
            {
                this.needDoCleanup = false;

                // If the server supports Public Folder Replicas, should set R1111Enabled in corresponding deployment ptfconfigure file as true.
                Site.Assert.Inconclusive("The exchange server doesn't support Public Folder Replicas.");
            }
        }

        /// <summary>
        /// This method is used to check whether MAPIHTTP transport is supported by SUT.
        /// </summary>
        /// <param name="isSupported">The transport is supported or not.</param>
        public void CheckMAPIHTTPTransportSupported(out bool isSupported)
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && !Common.IsRequirementEnabled(5100, this.Site))
            {
                isSupported = false;
                this.needDoCleanup = false;
                Site.Assert.Inconclusive("This case can not run on MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
            else
            {
                isSupported = true;
            }
        }

        /// <summary>
        /// Import new folders, or changes with conflict to existing folders, into the server replica.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="uploadContextHandleIndex">Upload context handle.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="properties">Properties to be set.</param>
        /// <param name="localFolderIdIndex">Local folder id index</param>
        /// <param name="folderIdIndex">The folder object id index.</param>
        /// <param name="conflictType">The conflict type will import.</param>
        /// <returns>Indicate the result of this ROP operation.</returns>
        public RopResult SynchronizationImportHierarchyChangeWithConflict(int serverId, int uploadContextHandleIndex, int parentFolderHandleIndex, Set<string> properties, int localFolderIdIndex, out int folderIdIndex, ConflictTypes conflictType)
        {
            // Initialize ROP data.
            uint uploadContextHandle = this.handleContainer[uploadContextHandleIndex];
            RopResult result = RopResult.InvalidParameter;
            folderIdIndex = -1;
            ulong folderId;

            // Construct ROP request.
            RopSynchronizationImportHierarchyChangeRequest synchronizationImportHierarchyChangeRequest;

            // Construct the RopSynchronizationImportHierarchyChange request.
            synchronizationImportHierarchyChangeRequest.RopId = 0x73;
            synchronizationImportHierarchyChangeRequest.LogonId = 0x00;
            synchronizationImportHierarchyChangeRequest.InputHandleIndex = 0x00;

            // Get the folder property value.
            TaggedPropertyValue[] taggedPropertyValueArrays = new TaggedPropertyValue[properties.Count];
            int index = 0;
            foreach (string tempProperty in properties)
            {
                taggedPropertyValueArrays[index++] = this.taggedPropertyValuesDictionary[tempProperty];
            }

            TaggedPropertyValue[] hierarchyValues = this.CreateSampleHierarchyValues(serverId, localFolderIdIndex, parentFolderHandleIndex, out folderId, conflictType);
            synchronizationImportHierarchyChangeRequest.HierarchyValueCount = (ushort)hierarchyValues.Length;
            synchronizationImportHierarchyChangeRequest.HierarchyValues = hierarchyValues;
            synchronizationImportHierarchyChangeRequest.PropertyValueCount = (ushort)taggedPropertyValueArrays.Length;
            synchronizationImportHierarchyChangeRequest.PropertyValues = taggedPropertyValueArrays;

            // Send request and get response.
            RopSynchronizationImportHierarchyChangeResponse synchronizationImportHierarchyChangeResponse = (RopSynchronizationImportHierarchyChangeResponse)this.Process(serverId, synchronizationImportHierarchyChangeRequest, uploadContextHandle);
            result = (RopResult)synchronizationImportHierarchyChangeResponse.ReturnValue;

            if (result == RopResult.Success)
            {
                if (!this.objectIdContainer.ContainsKey(localFolderIdIndex))
                {
                    folderIdIndex = AdapterHelper.GetObjectIdIndex();
                    this.objectIdContainer.Add(folderIdIndex, folderId);
                }

                // Verify ROP SynchronizationImportHierarchyChange
                this.VerifyRopSynchronizationImportHierarchyChange(synchronizationImportHierarchyChangeResponse);
            }

            return result;
        }

        /// <summary>
        /// Delete sub folders and messages
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        private void HardDeleteMessageAndSubfolders(int serverId)
        {
            int logonHandleIndex, inboxFolderIndex;

            // Reconnect the server
            bool resultDisconnect = this.oxcropsClient[serverId].Disconnect();
            Site.Assert.IsTrue(resultDisconnect, "Disconnecting server" + serverId + "should be successful.");
            this.Connect(serverId, Microsoft.Protocols.TestSuites.Common.ConnectionType.PrivateMailboxServer);
            this.Logon(serverId, LogonFlags.Private, out logonHandleIndex, out inboxFolderIndex);
            List<uint> soh;
            RopOpenFolderResponse res = this.OpenFolder(serverId, (uint)this.logonHandleOut, this.inboxFolderId, out soh);
            uint rootPublicFolderHandle = soh[res.OutputHandleIndex];

            uint folderHandle = rootPublicFolderHandle;
            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = 0x92,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };
            this.Process(serverId, hardDeleteMessagesAndSubfoldersRequest, folderHandle);
        }

        /// <summary>
        /// Initial the Variables.
        /// </summary>
        private void InitialVariables()
        {
            this.handleContainer = new Dictionary<int, uint>();
            this.objectIdContainer = new Dictionary<int, ulong>();
            this.streamBufferContainer = new Dictionary<int, byte[]>();
            this.propertyValuesSpecific = new Dictionary<string, byte[]>();
            this.icsStateContainer = new Dictionary<int, ICSStateData>();
            this.cnsetRead = new Dictionary<int, REPLGUID_IDSET>();
            this.cnsetSeen = new Dictionary<int, REPLGUID_IDSET>();
            this.cnsetSeenFAI = new Dictionary<int, REPLGUID_IDSET>();
            this.logonHandleOut = (uint)0;
            AdapterHelper.ClearIndex();
            this.InitialPropertyTagsDictionary();
            this.InitialTaggedPropertyValuesDictionary();
        }

        /// <summary>
        /// Send ROP request with single operation and get ROP response.
        /// </summary>
        /// <param name="serverId">A 32-bit signed integer represent the Identity of server.</param>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="inputHandle">Server object handle in request.</param>
        /// <returns>ROP response object.</returns>
        private IDeserializable Process(int serverId, ISerializable ropRequest, uint inputHandle)
        {
            List<ISerializable> inputBuffer = new List<ISerializable>
            {
                ropRequest
            };
            List<uint> requestSOH = new List<uint>();
            requestSOH.Add(inputHandle);
            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                requestSOH.Add(0xFFFFFFFF);
            }

            List<IDeserializable> responses = new List<IDeserializable>();
            List<List<uint>> responseSOHTable = new List<List<uint>>();
            byte[] rawData = null;
            this.ropResult = this.oxcropsClient[serverId].RopCall(inputBuffer, requestSOH, ref responses, ref responseSOHTable, ref rawData, 0x10008);

            if (this.ropResult == OxcRpcErrorCode.ECNone)
            {
                this.responseSOHs = responseSOHTable[0];
                this.VerifyMAPITransport();
            }

            if (responses.Count > 0)
            {
                return responses[0];
            }
            else
            {
                return null;
            }
        }

        #region Help functions for getting public folder handle.
        /// <summary>
        /// Create the PropertyTag sent to setColumn for GetHierarchyTable
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
        /// Set specific table column.
        /// </summary>
        /// <param name="serverID">The server ID.</param>
        /// <param name="inputHandle">Table handle.</param>
        /// <param name="setColumnsFlags">Set columns flags</param>
        /// <param name="propertyTags">Properties to be set.</param>
        /// <param name="serverObjectHandleTable">Server object handle table.</param>
        /// <returns>The response of RopSetColumns operation.</returns>
        private RopSetColumnsResponse SetColumns(
            int serverID,
            uint inputHandle,
            byte setColumnsFlags,
            PropertyTag[] propertyTags,
            out List<uint> serverObjectHandleTable)
        {
            RopSetColumnsRequest requestBuffer;

            // Construct ROP request.
            requestBuffer.RopId = 0x12;
            requestBuffer.LogonId = 0x00;
            requestBuffer.InputHandleIndex = 0x00;
            requestBuffer.SetColumnsFlags = setColumnsFlags;
            requestBuffer.PropertyTagCount = (ushort)propertyTags.Length;
            requestBuffer.PropertyTags = propertyTags;

            // Send the request.
            RopSetColumnsResponse rt = (RopSetColumnsResponse)this.Process(
                serverID,
                requestBuffer,
                inputHandle);
            serverObjectHandleTable = this.responseSOHs;

            return rt;
        }

        /// <summary>
        /// Open folder operation.
        /// </summary>
        /// <param name="serverID">server id.</param>
        /// <param name="inputHandle">Folder handle.</param>
        /// <param name="folderId">Folder id.</param>
        /// <param name="soh">Server object handle.</param>
        /// <returns>Response of RopOpenFolder operation.</returns>
        private RopOpenFolderResponse OpenFolder(
            int serverID,
            uint inputHandle,
            ulong folderId,
            out List<uint> soh)
        {
            RopOpenFolderRequest openFolderRequest;

            // Construct ROP request.
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = folderId;
            openFolderRequest.OpenModeFlags = 0x0;

            // Send request.
            RopOpenFolderResponse rt = (RopOpenFolderResponse)this.Process(
                serverID,
                openFolderRequest,
                inputHandle);
            soh = this.responseSOHs;
            return rt;
        }

        /// <summary>
        /// Get public folder handle.
        /// </summary>
        /// <param name="serverID">server id.</param>
        /// <param name="rootFolderID">Root folder handle.</param>
        /// <param name="name">Public folder name.</param>
        /// <param name="folderID">Folder id of public folder.</param>
        /// <param name="folderHandle">Folder handle of public folder.</param>
        private void GetPublicFolderHandle(
            int serverID,
            ulong rootFolderID,
            string name,
            out ulong folderID,
            out uint folderHandle)
        {
            List<uint> soh;
            RopOpenFolderResponse res = this.OpenFolder(
                        serverID,
                        this.logonHandleOut,
                        rootFolderID,
                        out soh);

            Debug.Assert(res.ReturnValue == 0, "This value specifies the status of the ROP.For this response, this field is set to 0x00000000.");
            uint rootFolderHandle = soh[res.OutputHandleIndex];

            RopGetHierarchyTableResponse getHierarchyTableResponse =
                this.GetHierarchyTable(
                serverID,
                rootFolderHandle,
                TableFlags.Depth,
                out this.responseSOHs);

            Debug.Assert(getHierarchyTableResponse.ReturnValue == 0, "This value specifies the status of the ROP.For this response, this field is set to 0x00000000.");
            uint tableHandle = this.responseSOHs[getHierarchyTableResponse.OutputHandleIndex];

            PropertyTag[] propertyTags = this.CreateHierarchyTablePropertyTagsForFolder();
            RopSetColumnsResponse setColRes = this.SetColumns(
                serverID,
                tableHandle,
                0,
                propertyTags,
                out this.responseSOHs);

            Debug.Assert(setColRes.ReturnValue == 0, "This value specifies the status of the ROP.For this response, this field is set to 0x00000000.");

            RopQueryRowsResponse queryRowRes = this.QueryRows(
                serverID,
                tableHandle,
                QueryRowsFlags.Advance,
                true,
                0xffff,
                out this.responseSOHs);

            Debug.Assert(queryRowRes.ReturnValue == 0, "This value specifies the status of the ROP.For this response, this field is set to 0x00000000.");

            // To avoid compile error.
            folderHandle = 0xffffffff;

            // To avoid compile error.
            folderID = 0xffffffffffffffff;

            bool hasFolder = false;

            for (int i = 0; i < queryRowRes.RowData.Count; i++)
            {
                byte[] buffer = queryRowRes.RowData.PropertyRows[i].PropertyValues[1].Value;
                string folderName = Encoding.Unicode.GetString(buffer);
                if (folderName == name + "\0")
                {
                    folderID = BitConverter.ToUInt64(
                        queryRowRes.RowData.PropertyRows[i].PropertyValues[0].Value,
                        0);
                    hasFolder = true;
                    break;
                }
            }

            if (!hasFolder)
            {
                AdapterHelper.Site.Assert.Fail("The given public folder name can't be found in the server.");
            }
            else
            {
                res = this.OpenFolder(
                    serverID,
                    rootFolderHandle,
                    folderID,
                    out soh);

                Debug.Assert(res.ReturnValue == 0, "This value specifies the status of the ROP. For this response, this field is set to 0x00000000.");
                folderHandle = soh[res.OutputHandleIndex];
            }
        }

        #endregion

        /// <summary>
        /// Insert state value into state dictionary.
        /// </summary>
        /// <param name="s">State object.</param>
        /// <returns>State value index in container.</returns>
        private int InsertStateDict(State s)
        {
            if (s != null && s.PropList != null && s.PropList.PropValues != null && s.PropList.PropValues.Count > 0)
            {
                int index = AdapterHelper.GetICSStateIndex();
                ICSStateData data = new ICSStateData();
                for (int i = 0; i < s.PropList.PropValues.Count; i++)
                {
                    PropValue value = s.PropList.PropValues[i];

                    VarPropTypePropValue varPropTypePropValue = value as VarPropTypePropValue;
                    AdapterHelper.Site.Assert.IsNotNull(varPropTypePropValue, "The VarPropTypePropValue should not be null.");

                    if (value.PropType == 0x0102 && varPropTypePropValue.PropInfo.PropID == 0x67D2)
                    {
                        // PidTagCnsetRead
                        data.PidTagCnsetRead = varPropTypePropValue.ValueArray;
                    }
                    else if (value.PropType == 0x0102 && varPropTypePropValue.PropInfo.PropID == 0x6796)
                    {
                        // PidTagCnsetSeen
                        data.PidTagCnsetSeen = varPropTypePropValue.ValueArray;
                    }
                    else if (value.PropType == 0x0102 && varPropTypePropValue.PropInfo.PropID == 0x67DA)
                    {
                        // PidTagCnsetSeenFAI
                        data.PidTagCnsetSeenFAI = varPropTypePropValue.ValueArray;
                    }
                    else if (value.PropType == 0x003 && varPropTypePropValue.PropInfo.PropID == 0x4017)
                    {
                        // PidTagIdsetGiven
                        data.PidTagIdsetGiven = varPropTypePropValue.ValueArray;
                    }
                    else
                    {
                        AdapterHelper.Site.Assert.Fail("Invalid property, its ID is {0} and type is {1}.", value.PropInfo.PropID.ToString("X4"), value.PropType.ToString("X4"));
                    }
                }

                this.icsStateContainer.Add(index, data);
                return index;
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("Invalid input state, ensure the state not null and have more than one PropValue in its PropList.");
                return -1;
            }
        }

        /// <summary>
        /// Indicate If IDSET is in dictionary.
        /// </summary>
        /// <param name="idset">IDSET value.</param>
        /// <param name="dict">IDSET dictionary.</param>
        /// <returns>IF id set is in it will return true.</returns>
        private bool IsInDict(REPLGUID_IDSET idset, Dictionary<int, REPLGUID_IDSET> dict)
        {
            AdapterHelper.Site.Assert.IsNotNull(idset, "The value of idset should not be null.");

            // If the dictionary does not contains anything,
            // the IDSET is not in the dictionary.
            if (dict.Values.Count <= 0)
            {
                return false;
            }
            else
            {
                // If any IDSET in the dictionary contains the IDSET
                // (the IDSET is a subset of any IDSET in the dictionary)
                foreach (REPLGUID_IDSET ids in dict.Values)
                {
                    if (ids.Contains(idset))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Indicate If IDSET is in dictionary.
        /// </summary>
        /// <param name="idset">IDSET value.</param>
        /// <param name="dict">IDSET dictionary.</param>
        /// <returns>IF id set is in it will return true.</returns>
        private bool HasSameCnset(REPLGUID_IDSET idset, Dictionary<int, REPLGUID_IDSET> dict)
        {
            if (!this.IsInDict(idset, dict))
            {
                return false;
            }
            else
            {
                foreach (REPLGUID_IDSET ids in dict.Values)
                {
                    if (ids.Contains(idset) && idset.Contains(ids))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Gets AbstractState structure.
        /// </summary>
        /// <param name="serverID">an int value,represents the id of a server</param>
        /// <param name="astateIndex">the index of a ICSStateData maintained 
        /// in the state dictionary(ICSStateContainer)</param>
        /// <param name="s">A state object.</param>
        /// <returns>The AbstractState corresponding to the state object.</returns>
        private AbstractState GetAbstractState(
            int serverID, int astateIndex, State s)
        {
            AbstractState astate = new AbstractState();

            // If CnsetRead which represents a list of change numbers contains in the state,
            // check whether the CnsetRead is newer than those contained in the dictionary _CnsetRead
            // if the CnsetRead is newer, add the CnsetRead to the dictionary _CnsetRead
            if (s.CnsetRead != null)
            {
                astate.IsNewCnsetReadPropertyChangeNumber
                    = !this.IsInDict(s.CnsetRead, this.cnsetRead);
                if (astate.IsNewCnsetReadPropertyChangeNumber)
                {
                    this.cnsetRead.Add(AdapterHelper.GetCnsetReadIndex(), s.CnsetRead);
                }
            }

            // If CnsetSeen which represents a list of change numbers contains in the state,
            // check whether the CnsetSeen is newer than those contained in the dictionary _CnsetSeen
            // if the CnsetSeen is newer, add the CnsetSeen to the dictionary _CnsetSeen
            if (s.CnsetSeen != null)
            {
                astate.IsNewCnsetSeenPropertyChangeNumber
                    = !this.IsInDict(s.CnsetSeen, this.cnsetSeen);

                // After the read state modified, the CN will not be changed in server 2007 and 2013.
                if (!astate.IsNewCnsetSeenPropertyChangeNumber)
                {
                    if (Common.IsRequirementEnabled(2246, this.Site))
                    {
                        if (this.cnsetSeen.Values.Count != 0)
                        {
                            astate.IsNewCnsetSeenPropertyChangeNumber
                                = this.HasSameCnset(s.CnsetSeen, this.cnsetSeen);
                        }
                    }
                }

                if (astate.IsNewCnsetSeenPropertyChangeNumber)
                {
                    this.cnsetSeen.Add(AdapterHelper.GetCnsetSeenIndex(), s.CnsetSeen);
                }
            }

            // If CnsetSeenFAI which represents a list of change numbers contains in the state,
            // check whether the CnsetSeenFAI is newer than those contained in the dictionary _CnsetSeenFAI
            // If the CnsetSeenFAI is newer, add the CnsetSeenFAI to the dictionary _CnsetSeenFAI
            if (s.CnsetSeenFAI != null)
            {
                astate.IsNewCnsetSeenFAIPropertyChangeNumber
                    = !this.IsInDict(s.CnsetSeenFAI, this.cnsetSeenFAI);
                if (astate.IsNewCnsetSeenFAIPropertyChangeNumber)
                {
                    this.cnsetSeenFAI.Add(AdapterHelper.GetCnsetSeenFAIIndex(), s.CnsetSeenFAI);
                }
            }

            // If IdsetGiven which convert to  a list of LongTermIds is not empty,
            // convert each LongTermId in the list of LongTermIds to message id,
            // and if the id is not contained in the objectIdContainer add the id 
            // to the IdSetGiven of the AbstractState.
            astate.IdSetGiven = new Set<int>();
            if (s.IdsetGiven != null)
            {
                List<LongTermId> lids = s.IdsetGiven.GetLongTermIds();
                List<ulong> ids = new List<ulong>();
                for (int i = 0; i < lids.Count; i++)
                {
                    ids.Add(this.GetIdFromLongTerm(serverID, lids[i]));
                }

                foreach (int key in this.objectIdContainer.Keys)
                {
                    if (ids.Contains(this.objectIdContainer[key]))
                    {
                        astate.IdSetGiven = astate.IdSetGiven.Add(key);
                    }
                }
            }

            // astateIndex is an incoming parameter.
            // Set the AbstractICSStateIndex of the AbstractICSState.
            astate.AbstractICSStateIndex = astateIndex;
            return astate;
        }

        /// <summary>
        /// Gets abstract deletion
        /// </summary>
        /// <param name="deletion">A deletion object.</param>
        /// <returns>The corresponding AbstractDeletion object.</returns>
        private AbstractDeletion GetAbstractDeletion(Deletions deletion)
        {
            AbstractDeletion ad = new AbstractDeletion
            {
                IdSetDeleted = new Modeling.Set<int>()
            };

            // If the deletion is not null change the AbstractDeletion's field,
            // else do nothing.
            if (deletion != null)
            {
                // If the deletion is not null the deletion is present.
                ad.IsDeletionPresent = true;

                // Check pidtags in the deletions property list.
                ad.IsPidTagIdsetExpiredExist
                    = deletion.PropList.HasPidTagIdsetExpired;
                ad.IsPidTagIdsetNoLongerInScopeExist
                    = deletion.PropList.HasPidTagIdsetNoLongerInScope;

                // Check IDSET in the deletion.
                // If the IDSET that can be got from PidTagIdsetDeleted value 
                // in the PropList of the deletions is contained in the objectIdContainer
                // add it to the IdsetDeleted field in the AbstractDeletion
                List<ulong> ids = deletion.IdsetDeleted;

                this.VerifyMetaTagIdsetDeletedProperty(this.hasExecuteSynchronizationImportDeletes, this.importDelFlag, ids, this.delObjId);
                this.hasExecuteSynchronizationImportDeletes = false;

                foreach (int key in this.objectIdContainer.Keys)
                {
                    if (ids.Contains(this.objectIdContainer[key]))
                    {
                        ad.IdSetDeleted = ad.IdSetDeleted.Add(key);
                    }
                }
            }

            return ad;
        }

        /// <summary>
        /// Get HierachySync data.
        /// </summary>
        /// <param name="serverId">server id.</param>
        /// <param name="hsync">Content sync object</param>
        /// <param name="stateIndex">Start index.</param>
        /// <returns>The AbstractFastTransferStream</returns>
        private AbstractFastTransferStream GetAbstractHierachySync(
            int serverId,
            HierarchySync hsync,
            int stateIndex)
        {
            AbstractFastTransferStream afts = new AbstractFastTransferStream();
            AbstractHierarchySync ahsync = new AbstractHierarchySync();
            afts.StreamType = FastTransferStreamType.hierarchySync;

            ahsync.FinalICSState = this.GetAbstractState(serverId, stateIndex, hsync.State);
            ahsync.AbstractDeletion = this.GetAbstractDeletion(hsync.Deletions);
            AbstractFolderChange fc = new AbstractFolderChange();

            // Get the FolderchangeInfo, FolderCount and IsParentFolderBeforeChild
            // of the AbstractHierarchySync.
            if (hsync.FolderChangeList != null && hsync.FolderChangeList.Count > 0)
            {
                AbstractFolderChange tmp = new AbstractFolderChange();

                // fields:IsPidTagFolderIdExist,IsPidTagParentFolderIdExist
                // IsPidTagParentSourceKeyValueZero,IsPidTagSourceKeyValueZero
                // need all folderChange's corresponding flag is true.
                fc = hsync.FolderChangeList[0].GetAbstractFolderChange();
                for (int i = 1; i < hsync.FolderChangeList.Count; i++)
                {
                    tmp = hsync.FolderChangeList[i].GetAbstractFolderChange();
                    fc.IsPidTagFolderIdExist =
                        fc.IsPidTagFolderIdExist && tmp.IsPidTagFolderIdExist;
                    fc.IsPidTagParentFolderIdExist =
                        fc.IsPidTagParentFolderIdExist && tmp.IsPidTagParentFolderIdExist;
                    fc.IsPidTagParentSourceKeyValueZero =
                        fc.IsPidTagParentSourceKeyValueZero && tmp.IsPidTagParentSourceKeyValueZero;
                    fc.IsPidTagSourceKeyValueZero =
                        fc.IsPidTagSourceKeyValueZero && tmp.IsPidTagSourceKeyValueZero;
                }

                ahsync.FolderchangeInfo = fc;

                // The count of folderChange.
                ahsync.FolderCount = hsync.FolderChangeList.Count;

                // If parents are before their children,
                // for a specific folder change, its ParentSourceKey
                // must be zero if its parent is the sync folder
                // or can be found before the position of the folder change.
                ahsync.IsParentFolderBeforeChild = hsync.IsParentBeforeChildren;
            }

            afts.AbstractHierarchySync = ahsync;
            return afts;
        }

        /// <summary>
        /// Get hierarchy table
        /// </summary>
        /// <param name="serverID">server id.</param>
        /// <param name="handle">Table handle</param>
        /// <param name="flag">A TableFlags value</param>
        /// <param name="soh">Server object handle.</param>
        /// <returns>The response of RopGetHierarchyTable operation.</returns>
        private RopGetHierarchyTableResponse GetHierarchyTable(
            int serverID,
            uint handle,
            TableFlags flag,
            out List<uint> soh)
        {
            RopGetHierarchyTableRequest getHierarchyTableRequest;

            // Construct ROP request.
            getHierarchyTableRequest.RopId = 0x04;
            getHierarchyTableRequest.LogonId = 0x00;
            getHierarchyTableRequest.InputHandleIndex = 0x00;
            getHierarchyTableRequest.OutputHandleIndex = 0x01;
            getHierarchyTableRequest.TableFlags = (byte)flag;

            // Send the request.
            RopGetHierarchyTableResponse rt = (RopGetHierarchyTableResponse)this.Process(
                serverID, getHierarchyTableRequest, handle);
            soh = this.responseSOHs;
            return rt;
        }

        /// <summary>
        /// Query table rows data.
        /// </summary>
        /// <param name="serverID">server id.</param>
        /// <param name="inputHandle">Table object handle.</param>
        /// <param name="queryRowsFlags">Query flags</param>
        /// <param name="forwardRead">Forward read.</param>
        /// <param name="rowCount">The number of rows.</param>
        /// <param name="serverObjectHandleTable">Server object handle table.</param>
        /// <returns>The response of RopQueryRows operation.</returns>
        private RopQueryRowsResponse QueryRows(
            int serverID,
            uint inputHandle,
            QueryRowsFlags queryRowsFlags,
            bool forwardRead,
            ushort rowCount,
            out List<uint> serverObjectHandleTable)
        {
            RopQueryRowsRequest queryRowsRequest;

            // Construct ROP request
            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = (byte)queryRowsFlags;
            queryRowsRequest.ForwardRead = (byte)(forwardRead ? 1 : 0);
            queryRowsRequest.RowCount = rowCount;

            // Send request.
            RopQueryRowsResponse rt = (RopQueryRowsResponse)this.Process(
                serverID,
                queryRowsRequest,
                inputHandle);
            serverObjectHandleTable = this.responseSOHs;
            return rt;
        }

        /// <summary>
        /// Generate fast transfer stream data.
        /// </summary>
        /// <param name="serverId">server id.</param>
        /// <param name="buffer">Input buffer data.</param>
        /// <returns>AbstractFastTransferStream object.</returns>
        private AbstractFastTransferStream GenerateAbstractFastTransferStream(
            int serverId,
            byte[] buffer)
        {
            using (FastTransferStream fs = new FastTransferStream(buffer, false))
            {
                AbstractFastTransferStream afts;

                // Record all property list generated while deserializing.
                SyntacticalBase.AllPropList = new List<PropList>();
                int icsStateIndex = -1;
                switch (this.streamType)
                {
                    case FastTransferStreamType.state:
                        State s = new State(fs);
                        this.VerifyMarkers(s);

                        // Insert state to the ICSStateContainer
                        icsStateIndex = this.InsertStateDict(s);
                        afts = new AbstractFastTransferStream();

                        // Get the AbstractState corresponding to the state from the fast transfer stream.
                        AbstractState astate = this.GetAbstractState(serverId, icsStateIndex, s);

                        // Set the AbstractState of the AbstractFastTransferStream.
                        afts.AbstractState = new AbstractState
                        {
                            AbstractICSStateIndex = icsStateIndex,
                            IdSetGiven = astate.IdSetGiven
                        };

                        //// Other fields of the AbstractState of the AbstractFastTransferStream do not need.

                        afts.StreamType = FastTransferStreamType.state;
                        return afts;
                    case FastTransferStreamType.attachmentContent:
                        AttachmentContent att = new AttachmentContent(fs);
                        this.VerifyMarkers(att);
                        return att.GetAbstractFastTransferStream();
                    case FastTransferStreamType.contentsSync:
                        ContentsSync cs = new ContentsSync(fs);
                        this.VerifyMarkers(cs);

                        // Insert the state of the contentsSync to the ICSStateContainer
                        icsStateIndex = this.InsertStateDict(cs.State);
                        afts = this.GetAbstractContentSync(
                            serverId, icsStateIndex, cs);
                        return afts;
                    case FastTransferStreamType.folderContent:
                        FolderContent fc = new FolderContent(fs);
                        this.VerifyMarkers(fc);
                        this.VerifyFolderReplicaInfoStructure(fc);

                        afts = fc.GetAbstractFastTransferStream();

                        if (!this.existNoPermissionFolder)
                        {
                            afts.AbstractFolderContent.IsNoPermissionObjNotOut = false;
                        }

                        this.VerifyMetaProperty(afts);
                        return afts;
                    case FastTransferStreamType.hierarchySync:
                        HierarchySync hs = new HierarchySync(fs);
                        this.VerifyMarkers(hs);
                        if (hs.FolderChangeList.Count > 0)
                        {
                            PropValue property = hs.FolderChangeList[hs.FolderChangeList.Count - 1].PropList.PropValues.Find(p => p.PropInfo.PropID == 0x65e3);
                            if (property != null)
                            {
                                this.lastConflictInfo.PCLXFromServer = ((VarPropTypePropValue)property).ValueArray;
                                this.VerifyPidTagPredecessorChangeList();
                            }
                        }

                        // Insert state to the ICSStateContainer
                        icsStateIndex = this.InsertStateDict(hs.State);
                        afts = this.GetAbstractHierachySync(serverId, hs, icsStateIndex);
                        return afts;
                    case FastTransferStreamType.MessageContent:
                        MessageContent mc = new MessageContent(fs);
                        for (int i = 0; i < mc.PropList.PropValues.Count; i++)
                        {
                            PropValue propValue = mc.PropList.PropValues[i];

                            if (propValue.PropType == 0x84b0)
                            {
                                CodePage codePage = new CodePage();
                                codePage.Deserialize(propValue.PropType);

                                this.VerifyCodePageProperty(codePage);
                            }
                        }

                        this.VerifyMarkers(mc);
                        afts = mc.GetAbstractFastTransferStream();
                        this.VerifyMetaProperty(afts);
                        return afts;
                    case FastTransferStreamType.MessageList:
                        MessageList ml = new MessageList(fs);
                        afts = ml.GetAbstractFastTransferStream();
                        this.VerifyMetaProperty(afts);
                        this.VerifyMarkers(ml);
                        return afts;
                    case FastTransferStreamType.TopFolder:
                        TopFolder tf = new TopFolder(fs);
                        this.VerifyMarkers(tf);
                        afts = tf.GetAbstractFastTransferStream();
                        this.VerifyMetaProperty(afts);
                        return afts;
                    default:
                        AdapterHelper.Site.Assert.Fail("The stream type is invalid, its value is:{0}.", this.streamType);
                        return new AbstractFastTransferStream();
                }
            }
        }

        #region Helper
        /// <summary>
        /// Add Recipient to specific message.
        /// </summary>
        /// <param name="serverId">Server id.</param>
        /// <param name="messageHandle">Message handle.</param>
        private void AddRecipient(int serverId, uint messageHandle)
        {
            RopModifyRecipientsRequest modifyRecipientRequest = new RopModifyRecipientsRequest
            {
                RopId = 0x0E,
                LogonId = 0x00,
                InputHandleIndex = 0x00
            };
            PropertyTag[] tags = null;
            ModifyRecipientRow[] recipientRow = null;
            this.CreateSampleRecipientColumnsAndRecipientRows(out tags, out recipientRow);
            modifyRecipientRequest.RecipientColumns = tags;
            modifyRecipientRequest.RecipientRows = recipientRow;
            modifyRecipientRequest.RowCount = (ushort)recipientRow.Length;
            modifyRecipientRequest.ColumnCount = (ushort)tags.Length;
            this.Process(serverId, modifyRecipientRequest, messageHandle);
        }

        /// <summary>
        /// Generate recipientColumns value and recipientRows value.
        /// </summary>
        /// <param name="recipientColumns">recipientColumns value.</param>
        /// <param name="recipientRows">recipientRows value.</param>
        private void CreateSampleRecipientColumnsAndRecipientRows(out PropertyTag[] recipientColumns, out ModifyRecipientRow[] recipientRows)
        {
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            PropertyTag tag;

            // PidTagObjectType
            tag.PropertyId = 0x0ffe;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[0] = tag;

            // PidTagDisplayType
            tag.PropertyId = 0x3900;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[1] = tag;

            // PidTagAddressBookDisplayNamePrintable
            tag.PropertyId = 0x39ff;

            // PtypString
            tag.PropertyType = 0x001f;
            sampleRecipientColumns[2] = tag;

            // PidTagSmtpAddress
            tag.PropertyId = 0x39fe;

            // PtypString
            tag.PropertyType = 0x001f;
            sampleRecipientColumns[3] = tag;

            // PidTagSendInternetEncoding
            tag.PropertyId = 0x3a71;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[4] = tag;

            // PidTagDisplayTypeEx
            tag.PropertyId = 0x3905;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[5] = tag;

            // PidTagRecipientDisplayName
            tag.PropertyId = 0x5ff6;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[6] = tag;

            // PidTagRecipientFlags
            tag.PropertyId = 0x5ffd;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[7] = tag;

            // PidTagRecipientTrackStatus
            tag.PropertyId = 0x5fff;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[8] = tag;

            // PidTagRecipientResourceState
            tag.PropertyId = 0x5fde;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[9] = tag;

            // PidTagRecipientOrder
            tag.PropertyId = 0x5fdf;

            // PtypInteger32
            tag.PropertyType = 0x0003;
            sampleRecipientColumns[10] = tag;

            // PidTagRecipientEntryId
            tag.PropertyId = 0x5ff7;

            // PtypBinary
            tag.PropertyType = 0x0102;
            sampleRecipientColumns[11] = tag;

            recipientColumns = sampleRecipientColumns;

            PropertyValue[] propertyValueArray = new PropertyValue[12];
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006); // PidTagObjectType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000); // PidTagDisplayType
            propertyValueArray[2].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "\0"); // PidTa7BitDisplayName
            propertyValueArray[3].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0"); // PidTagSmtpAddress
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000); // PidTagSendInternetEncoding
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000); // PidTagDisplayTypeEx
            propertyValueArray[6].Value = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "\0"); // PidTagRecipientDisplayName
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001); // PidTagRecipientFlags
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientTrackStatus
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientResourceState
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientOrder

            // EntryId of administrator 
            byte[] sampleData = ConvertStringToBytes(@"0x79,0x00,0x00,0x00,0x00,0x00,0xDC,0xA7,0x40,0xC8,0xC0,0x42
                                                           ,0x10,0x1A,0xB4,0xB9,0x08,0x00,0x2B,0x2F,0xE1,0x82
                                                           ,0x01,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x2F,0x6F
                                                           ,0x3D,0x63,0x6F,0x6E,0x74,0x6F,0x73,0x6F,0x2F,0x6F
                                                           ,0x75,0x3D,0x45,0x78,0x63,0x68,0x61,0x6E,0x67,0x65
                                                           ,0x20,0x41,0x64,0x6D,0x69,0x6E,0x69,0x73,0x74,0x72
                                                           ,0x61,0x74,0x69,0x76,0x65,0x20,0x47,0x72,0x6F,0x75
                                                           ,0x70,0x20,0x28,0x46,0x59,0x44,0x49,0x42,0x4F,0x48
                                                           ,0x46,0x32,0x33,0x53,0x50,0x44,0x4C,0x54,0x29,0x2F
                                                           ,0x63,0x6E,0x3D,0x52,0x65,0x63,0x69,0x70,0x69,0x65
                                                           ,0x6E,0x74,0x73,0x2F,0x63,0x6E,0x3D,0x41,0x64,0x6D
                                                           ,0x69,0x6E,0x69,0x73,0x74,0x72,0x61,0x74,0x6F,0x72
                                                           ,0x00");
            propertyValueArray[11].Value = sampleData;

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = 0x00,
                PropertyValues = propertyValues
            };

            // For propertyRow.Flag.
            int tempLengthForTest = 1;
            foreach (PropertyValue pv in propertyRow.PropertyValues)
            {
                tempLengthForTest = tempLengthForTest + pv.Value.Length;
            }

            RecipientRow recipientRow = new RecipientRow
            {
                RecipientFlags = 0x065B,

                // Present because D is Set.
                DisplayName = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "\0"),
                EmailAddress = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "@" +
                                              Common.GetConfigurationPropertyValue("Domain", this.Site) + "\0"),
                SimpleDisplayName = Encoding.Unicode.GetBytes(Common.GetConfigurationPropertyValue("AdminUserName", this.Site) + "\0"),

                // Matches ColummnCount.
                RecipientColumnCount = 0x000C,
                RecipientProperties = propertyRow
            };

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                // Primary Recipient.
                RowId = 0x00000000,
                RecipientType = 0x01,
                RecipientRowSize = (ushort)recipientRow.Size(),

                // Bytes in the following RecipientRow.
                RecptRow = recipientRow.Serialize()
            };

            ModifyRecipientRow[] sampleModifyRecipientRows = new ModifyRecipientRow[1];
            sampleModifyRecipientRows[0] = modifyRecipientRow;
            recipientRows = sampleModifyRecipientRows;
        }

        /// <summary>
        /// Get changeNmuber of a specific message object.
        /// </summary>
        /// <param name="serverId">Server id.</param>
        /// <param name="messageId">ID of the message object.</param>
        /// <param name="folderId">ID of the message's parent folder.</param>
        /// <param name="folderHandle">Handle of the message's parent folder</param>
        /// <returns>Change number.</returns>
        private byte[] GetChangeNumber(int serverId, ulong messageId, ulong folderId, uint folderHandle)
        {
            // Open message to get message handle.
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest();
            RopOpenMessageResponse openMessageResponse = new RopOpenMessageResponse();
            openMessageRequest.CodePageId = 0x0fff;
            openMessageRequest.FolderId = folderId;
            openMessageRequest.InputHandleIndex = 0x00;
            openMessageRequest.LogonId = 0x00;
            openMessageRequest.MessageId = messageId;
            openMessageRequest.OpenModeFlags = 0x01;
            openMessageRequest.OutputHandleIndex = 0x01;
            openMessageRequest.RopId = 0x03;
            openMessageResponse = (RopOpenMessageResponse)this.Process(serverId, openMessageRequest, folderHandle);
            uint handle = this.responseSOHs[openMessageResponse.OutputHandleIndex];

            // Get message PidTagChangeKey value.
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = new RopGetPropertiesSpecificResponse();
            PropertyTag[] propertyTags = new PropertyTag[1];
            PropertyTag tag = new PropertyTag(0x65E2, 0x0102);
            propertyTags[0] = tag;

            getPropertiesSpecificRequest.RopId = 0x07;
            getPropertiesSpecificRequest.LogonId = 0x00;
            getPropertiesSpecificRequest.InputHandleIndex = 0x00;

            // Set PropertySizeLimit, which specifies the maximum size allowed for a property value returned,
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;
            PropertyTag[] tagArray = propertyTags;
            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = tagArray;

            // Send the RopGetPropertiesSpecific request
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.Process(serverId, getPropertiesSpecificRequest, handle);

            if (getPropertiesSpecificResponse.ReturnValue == 0x0)
            {
                byte[] tempChangeNumber = new byte[getPropertiesSpecificResponse.RowData.PropertyValues[0].Value.Length - 2];
                Array.Copy(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, 2, tempChangeNumber, 0, getPropertiesSpecificResponse.RowData.PropertyValues[0].Value.Length - 2);
                return tempChangeNumber;
            }
            else
            {
                return new byte[0];
            }
        }

        /// <summary>
        /// Get longTermId bytes.
        /// </summary>
        /// <param name="serverId">The server ID</param>
        /// <param name="objectId">The id need to be converted</param>
        /// <returns>The converted longTermId from id.</returns>
        private byte[] GetLongTermIdByte(int serverId, ulong objectId)
        {
            RopLongTermIdFromIdRequest ropLongTermIdFromIdRequest = new RopLongTermIdFromIdRequest();
            RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse = new RopLongTermIdFromIdResponse();

            // Construct the RopLongTermIdFromId request.
            ropLongTermIdFromIdRequest.RopId = 0x43;
            ropLongTermIdFromIdRequest.LogonId = 0x00;
            ropLongTermIdFromIdRequest.InputHandleIndex = 0x00;
            ropLongTermIdFromIdRequest.ObjectId = objectId;

            // Send the RopLongTermIdFromId request to convert the short-term ID into a long-term ID.
            ropLongTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.Process(serverId, ropLongTermIdFromIdRequest, this.logonHandleOut);
            byte[] longTermByte = new byte[ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid.Length + ropLongTermIdFromIdResponse.LongTermId.GlobalCounter.Length];
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid, 0, longTermByte, 0, ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid.Length);
            Array.Copy(ropLongTermIdFromIdResponse.LongTermId.GlobalCounter, 0, longTermByte, ropLongTermIdFromIdResponse.LongTermId.DatabaseGuid.Length, ropLongTermIdFromIdResponse.LongTermId.GlobalCounter.Length);
            return longTermByte;
        }

        /// <summary>
        /// Convert longTermId to id.
        /// </summary>
        /// <param name="serverId">server id.</param>
        /// <param name="longTermId">longTerm id.</param>
        /// <returns>Converted short term id.</returns>
        private ulong GetIdFromLongTerm(int serverId, LongTermId longTermId)
        {
            RopIdFromLongTermIdRequest ropIdFromLongTermRequest = new RopIdFromLongTermIdRequest();
            RopIdFromLongTermIdResponse ropIdFromLongTermResponse = new RopIdFromLongTermIdResponse();
            ropIdFromLongTermRequest.RopId = 0x44;
            ropIdFromLongTermRequest.LogonId = 0x00;
            ropIdFromLongTermRequest.InputHandleIndex = 0x00;
            ropIdFromLongTermRequest.LongTermId = longTermId;
            ropIdFromLongTermResponse = (RopIdFromLongTermIdResponse)this.Process(serverId, ropIdFromLongTermRequest, this.logonHandleOut);
            return ropIdFromLongTermResponse.ObjectId;
        }

        /// <summary>
        /// Initialize the taggedPropertyValuesDictionary.
        /// </summary>
        private void InitialTaggedPropertyValuesDictionary()
        {
            TaggedPropertyValue propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x3001, 0x001F),
                Value = Encoding.Unicode.GetBytes("TestDisplayName" + "\0")
            };

            // PidTagDisplayName
            this.taggedPropertyValuesDictionary.Add("PidTagDisplayName", propertyvalue);

            // PidTagMessageClass
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x001A, 0x001F),
                Value = Encoding.Unicode.GetBytes("IPM.Note" + "\0")
            };
            this.taggedPropertyValuesDictionary.Add("PidTagMessageClass", propertyvalue);

            // PidTagAccess
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x0FF4, 0x0003)
            };
            int propertyValue = 0x02;
            propertyvalue.Value = BitConverter.GetBytes(propertyValue);
            this.taggedPropertyValuesDictionary.Add("PidTagAccess", propertyvalue);

            // PidTagAccessLevel
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x0FF7, 0x0003)
            };

            propertyValue = 0x00;
            propertyvalue.Value = BitConverter.GetBytes(propertyValue);
            this.taggedPropertyValuesDictionary.Add("PidTagAccessLevel", propertyvalue);

            // PidTagCreationTime
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x3007, 0x0040),
                Value = BitConverter.GetBytes(DateTime.UtcNow.ToFileTimeUtc())
            };
            this.taggedPropertyValuesDictionary.Add("PidTagCreationTime", propertyvalue);

            // PidTagBody
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x1000, 0x001F),
                Value = Encoding.Unicode.GetBytes("hello" + "\0")
            };

            this.taggedPropertyValuesDictionary.Add("PidTagBody", propertyvalue);

            // PidTagFolderType
            propertyvalue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x3601, 0x0003),
                Value = new byte[]
                {
                    0x01, 0x00, 0x00, 0x00
                }
            };

            // FOLDER_GENERIC
            this.taggedPropertyValuesDictionary.Add("PidTagFolderType", propertyvalue);
        }

        /// <summary>
        /// Initialize the propertyTagsDictionary.
        /// </summary>
        private void InitialPropertyTagsDictionary()
        {
            PropertyTag property = new PropertyTag
            {
                PropertyId = 0x0FF4,
                PropertyType = 0x0003
            };

            // PidTagAccess
            this.propertyTagsDictionary.Add("PidTagAccess", property);

            // PidTagAccessLevel
            property.PropertyId = 0x0FF7;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagAccessLevel", property);

            // PidTagChangeKey
            property.PropertyId = 0x65E2;
            property.PropertyType = 0x0102;
            this.propertyTagsDictionary.Add("PidTagChangeKey", property);

            // PidTagCreationTime
            property.PropertyId = 0x3007;
            property.PropertyType = 0x0040;
            this.propertyTagsDictionary.Add("PidTagCreationTime", property);

            // PidTagLastModificationTime
            property.PropertyId = 0x3008;
            property.PropertyType = 0x0040;
            this.propertyTagsDictionary.Add("PidTagLastModificationTime", property);

            // PidTagAttachLongFilename
            property.PropertyId = 0x3707;
            property.PropertyType = 0x001F;
            this.propertyTagsDictionary.Add("PidTagAttachLongFilename", property);

            // PidTagFolderId
            property.PropertyId = 0x6748;
            property.PropertyType = 0x0014;
            this.propertyTagsDictionary.Add("PidTagFolderId", property);

            // PidTagDisplayName
            property.PropertyId = 0x3001;
            property.PropertyType = 0x001F;
            this.propertyTagsDictionary.Add("PidTagDisplayName", property);

            // PidTagComment
            property.PropertyId = 0x3004;
            property.PropertyType = 0x001F;
            this.propertyTagsDictionary.Add("PidTagComment", property);

            // PidTagContainerContents
            property.PropertyId = 0x360f;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagContainerContents", property);

            // PidTagFolderAssociatedContents
            property.PropertyId = 0x3610;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagFolderAssociatedContents", property);

            // PidTagContainerHierarchy
            property.PropertyId = 0x360E;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagContainerHierarchy", property);

            // PidTagMessageRecipients
            property.PropertyId = 0x0E12;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagMessageRecipients", property);

            // PidTagMessageAttachments
            property.PropertyId = 0x0E13;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagMessageAttachments", property);

            // PidTagAttachDataObject
            property.PropertyId = 0x3701;
            property.PropertyType = 0x000d;
            this.propertyTagsDictionary.Add("PidTagAttachDataObject", property);

            // PidTagSourceKey
            property.PropertyId = 0x65E0;
            property.PropertyType = 0x0102;
            this.propertyTagsDictionary.Add("PidTagSourceKey", property);

            // PidTagPredecessorChangeList
            property.PropertyId = 0x65E3;
            property.PropertyType = 0x0102;
            this.propertyTagsDictionary.Add("PidTagPredecessorChangeList", property);

            // PidTagIdsetGivenPtyp by the Binary type
            property.PropertyId = 0x4017;
            property.PropertyType = 0x0102;
            this.propertyTagsDictionary.Add("PidTagIdsetGivenPtypBinary", property);

            // PidTagIdsetGivenPtyp by the Integer32 type
            property.PropertyId = 0x4017;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagIdsetGivenPtypInteger32", property);

            // PidTagAttachNumber 
            property.PropertyId = 0x0E21;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagAttachNumber", property);

            // PidTagBody
            property.PropertyId = 0x1000;
            property.PropertyType = 0x001F;
            this.propertyTagsDictionary.Add("PidTagBody", property);

            // PidTagEcWarning
            property.PropertyId = 0x400F;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagEcWarning", property);

            // PidTagResolveMethod
            property.PropertyId = 0x3FE7;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagResolveMethod", property);

            // PidTagPolicyTag
            property.PropertyId = 0x3019;
            property.PropertyType = 0x0102;
            this.propertyTagsDictionary.Add("PidTagPolicyTag", property);

            // PidTagLocalCommitTimeMax
            property.PropertyId = 0x670a;
            property.PropertyType = 0x0040;
            this.propertyTagsDictionary.Add("PidTagLocalCommitTimeMax", property);

            // PidTagDeletedCountTotal
            property.PropertyId = 0x670b;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagDeletedCountTotal", property);

            // PidTagLocalCommitTimeMax
            property.PropertyId = 0x67a4;
            property.PropertyType = 0x0014;
            this.propertyTagsDictionary.Add("PidTagChangeNumber", property);

            // PidTagMessageFlags
            property.PropertyId = 0x0e07;
            property.PropertyType = 0x0003;
            this.propertyTagsDictionary.Add("PidTagMessageFlags", property);
        }

        /// <summary>
        /// Create values for SynchronizationImportHierarchyChange.
        /// </summary>
        /// <param name="serverId">server id.</param>
        /// <param name="folderIdIndex">Folder id index.</param>
        /// <param name="parentFolderHandleIndex">Parent folder handle index.</param>
        /// <param name="folderId">Folder id.</param>
        /// <param name="conflictType">Conflict type </param>
        /// <returns>Return the constructed hierarchy values.</returns>
        private TaggedPropertyValue[] CreateSampleHierarchyValues(int serverId, int folderIdIndex, int parentFolderHandleIndex, out ulong folderId, ConflictTypes conflictType)
        {
            folderId = 0;
            TaggedPropertyValue[] hierarchyValues = new TaggedPropertyValue[6];
            TaggedPropertyValue propertyValue = new TaggedPropertyValue();
            byte[] currentChangeNumber = new byte[6];

            // PidTagParentSourceKey equal to the PidTagSourceKey of the parent folder
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x65E1, 0x0102)
            };

            Sequence<string> propertyTag = new Sequence<string>("PidTagSourceKey");
            this.GetPropertiesSpecific(serverId, parentFolderHandleIndex, propertyTag);
            propertyValue.Value = this.propertyValuesSpecific["PidTagSourceKey"];
            hierarchyValues[0] = propertyValue;

            if (conflictType == ConflictTypes.NONE)
            {
                // Whether the hierarchy already exists.
                if (this.objectIdContainer.ContainsKey(folderIdIndex))
                {
                    // Construct PidTagSourceKey value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E0, 0x0102)
                    };

                    byte[] sampleForPidTagSourceKey = this.propertyValuesSpecific["PidTagSourceKey"];
                    propertyValue.Value = sampleForPidTagSourceKey;
                    hierarchyValues[1] = propertyValue;

                    folderId = this.objectIdContainer[folderIdIndex];

                    // Construct PidTagLastModificationTime value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x3008, 0x0040)
                    };

                    byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                    propertyValue.Value = sampleForPidTagLastModificationTime;
                    hierarchyValues[2] = propertyValue;

                    // Construct PidTagChangeKey value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E2, 0x0102)
                    };

                    byte[] sampleForPidTagChangeKey = this.propertyValuesSpecific["PidTagChangeKey"];

                    // Modify ChangeKey to generate a new ChangeKey.
                    byte[] temp = new byte[4];
                    Array.Copy(sampleForPidTagChangeKey, sampleForPidTagChangeKey.Length - 4, temp, 0, 4);
                    Array.Reverse(temp);
                    int add = BitConverter.ToInt32(temp, 0);

                    // Modify the changeNumber to a newer version.
                    add++;
                    temp = BitConverter.GetBytes(add);
                    Array.Reverse(temp);
                    Array.Copy(temp, 0, sampleForPidTagChangeKey, sampleForPidTagChangeKey.Length - 4, 4);
                    propertyValue.Value = sampleForPidTagChangeKey;
                    hierarchyValues[3] = propertyValue;

                    // Construct PidTagPredecessorChangeList value
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E3, 0x0102)
                    };

                    byte[] sampleForPidTagPredecessorChangeList = new byte[this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length + sampleForPidTagChangeKey.Length];
                    Array.Copy(this.propertyValuesSpecific["PidTagPredecessorChangeList"], 0, sampleForPidTagPredecessorChangeList, 0, this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length);
                    Array.Copy(sampleForPidTagChangeKey, 0, sampleForPidTagPredecessorChangeList, this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length, sampleForPidTagChangeKey.Length);
                    propertyValue.Value = sampleForPidTagPredecessorChangeList;
                    hierarchyValues[4] = propertyValue;
                }
                else if (folderIdIndex == 0)
                {
                    // Get a local id from the id set given by server.
                    byte[] sub = new byte[4];
                    byte[] currentLocalId = this.localId;
                    Array.Copy(currentLocalId, 2, sub, 0, currentLocalId.Length - 2);
                    Array.Reverse(sub);
                    int totalNum = BitConverter.ToInt32(sub, 0);
                    totalNum += this.localIdOffSet;
                    sub = BitConverter.GetBytes(totalNum);
                    Array.Reverse(sub);
                    Array.Copy(sub, 0, currentLocalId, 2, currentLocalId.Length - 2);
                    this.localIdOffSet++;

                    // Construct PidTagSourceKey value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E0, 0x0102)
                    };

                    byte[] sampleForPidTagSourceKey = new byte[24];
                    int length = 0;
                    int index = 0;

                    // The combination of first two bytes indicates the length of value field.
                    length = (short)(this.serverReplicaGuid.ToByteArray().Length + currentLocalId.Length);
                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagSourceKey, 0, sizeof(short));
                    index += sizeof(short);
                    Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, sampleForPidTagSourceKey, index, this.serverReplicaGuid.ToByteArray().Length);
                    index += this.serverReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentLocalId, 0, sampleForPidTagSourceKey, index, currentLocalId.Length);
                    Array.Copy(currentLocalId, 0, currentChangeNumber, 0, currentLocalId.Length);
                    propertyValue.Value = sampleForPidTagSourceKey;
                    hierarchyValues[1] = propertyValue;

                    // Convert long term id to id .
                    LongTermId longTermId = new LongTermId
                    {
                        DatabaseGuid = this.serverReplicaGuid.ToByteArray(),
                        GlobalCounter = currentLocalId
                    };
                    folderId = this.GetIdFromLongTerm(serverId, longTermId);

                    // Construct PidTagLastModificationTime value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x3008, 0x0040)
                    };

                    byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                    propertyValue.Value = sampleForPidTagLastModificationTime;
                    hierarchyValues[2] = propertyValue;

                    // Construct PidTagChangeKey value.
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E2, 0x0102)
                    };

                    byte[] sampleForPidTagChangeKey = new byte[24];
                    index = 0;
                    length = 0;
                    length = (short)GidLength;

                    // The combination of first two bytes indicates the length of value field.
                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagChangeKey, 0, sizeof(short));
                    index += sizeof(short);
                    Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagChangeKey, index, this.localReplicaGuid.ToByteArray().Length);
                    index += this.localReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentChangeNumber, 0, sampleForPidTagChangeKey, index, currentChangeNumber.Length);
                    propertyValue.Value = sampleForPidTagChangeKey;
                    hierarchyValues[3] = propertyValue;

                    // Construct PidTagPredecessorChangeList value
                    propertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = new PropertyTag(0x65E3, 0x0102)
                    };

                    byte[] sampleForPidTagPredecessorChangeList = new byte[25];
                    index = 0;
                    length = 0;
                    length = (short)(GidLength + 1);
                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                    index += sizeof(short);
                    sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + currentChangeNumber.Length); // 16
                    index += 1;
                    Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                    index += this.localReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentChangeNumber, 0, sampleForPidTagPredecessorChangeList, index, currentChangeNumber.Length);
                    propertyValue.Value = sampleForPidTagPredecessorChangeList;
                    hierarchyValues[4] = propertyValue;
                }
            }
            else
            {
                // Get a local id from the id set given by server.
                byte[] sub = new byte[4];
                byte[] currentLocalId = this.localId;
                Array.Copy(currentLocalId, 2, sub, 0, currentLocalId.Length - 2);
                Array.Reverse(sub);
                int totalNum = BitConverter.ToInt32(sub, 0);
                totalNum += this.localIdOffSet;
                sub = BitConverter.GetBytes(totalNum);
                Array.Reverse(sub);
                Array.Copy(sub, 0, currentLocalId, 2, currentLocalId.Length - 2);
                this.localIdOffSet++;

                // Construct PidTagSourceKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E0, 0x0102)
                };

                byte[] sampleForPidTagSourceKey = new byte[24];
                int length = 0;
                int index = 0;

                // The combination of first two bytes indicates the length of value field.
                length = (short)(this.serverReplicaGuid.ToByteArray().Length + currentLocalId.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagSourceKey, 0, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, sampleForPidTagSourceKey, index, this.serverReplicaGuid.ToByteArray().Length);
                index += this.serverReplicaGuid.ToByteArray().Length;
                Array.Copy(currentLocalId, 0, sampleForPidTagSourceKey, index, currentLocalId.Length);
                Array.Copy(currentLocalId, 0, currentChangeNumber, 0, currentLocalId.Length);
                propertyValue.Value = sampleForPidTagSourceKey;
                hierarchyValues[1] = propertyValue;

                // Convert long term id to id .
                LongTermId longTermId = new LongTermId
                {
                    DatabaseGuid = this.serverReplicaGuid.ToByteArray(),
                    GlobalCounter = currentLocalId
                };
                folderId = this.GetIdFromLongTerm(serverId, longTermId);

                // Construct PidTagLastModificationTime value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x3008, 0x0040)
                };

                byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                propertyValue.Value = sampleForPidTagLastModificationTime;
                hierarchyValues[2] = propertyValue;

                // Construct PidTagChangeKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E2, 0x0102)
                };

                byte[] sampleForPidTagChangeKey = new byte[24];
                index = 0;
                length = 0;
                length = (short)GidLength;

                // The combination of first two bytes indicates the length of value field.
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagChangeKey, 0, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagChangeKey, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(currentChangeNumber, 0, sampleForPidTagChangeKey, index, currentChangeNumber.Length);
                propertyValue.Value = sampleForPidTagChangeKey;
                hierarchyValues[3] = propertyValue;

                // Construct PidTagPredecessorChangeList value
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E3, 0x0102)
                };

                byte[] sampleForPidTagPredecessorChangeList;

                if (conflictType == ConflictTypes.AINCLUDEB)
                {
                    sampleForPidTagPredecessorChangeList = new byte[25];
                    index = 0;
                    length = 0;
                    length = (short)(GidLength + 1);

                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                    index += sizeof(short);
                    sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + currentChangeNumber.Length); // 16
                    index += 1;
                    Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                    index += this.localReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentChangeNumber, 0, sampleForPidTagPredecessorChangeList, index, currentChangeNumber.Length);

                    this.lastConflictInfo.PCLB = sampleForPidTagPredecessorChangeList;
                    sampleForPidTagPredecessorChangeList[24] += 1;
                    this.lastConflictInfo.PCLA = sampleForPidTagPredecessorChangeList;
                    this.lastConflictInfo.PCLXFromMath = sampleForPidTagPredecessorChangeList;
                    this.lastConflictInfo.DetectedResult = conflictType;

                    propertyValue.Value = sampleForPidTagPredecessorChangeList;
                    hierarchyValues[4] = propertyValue;
                }
                else if (conflictType == ConflictTypes.BINCLUDEA)
                {
                    sampleForPidTagPredecessorChangeList = new byte[25];
                    index = 0;
                    length = 0;
                    length = (short)(GidLength + 1);

                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                    index += sizeof(short);
                    sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + currentChangeNumber.Length); // 16
                    index += 1;
                    Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                    index += this.localReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentChangeNumber, 0, sampleForPidTagPredecessorChangeList, index, currentChangeNumber.Length);

                    this.lastConflictInfo.PCLB = sampleForPidTagPredecessorChangeList;
                    this.lastConflictInfo.PCLXFromMath = sampleForPidTagPredecessorChangeList;
                    sampleForPidTagPredecessorChangeList[24] -= 1;
                    this.lastConflictInfo.PCLA = sampleForPidTagPredecessorChangeList;
                    this.lastConflictInfo.DetectedResult = conflictType;

                    propertyValue.Value = sampleForPidTagPredecessorChangeList;
                    hierarchyValues[4] = propertyValue;
                }

                if (conflictType == ConflictTypes.CONFLICT)
                {
                    // If the versions from client and server are in conflict, the value of PidTagPredecessorChangeList property in client should not include 
                    // the one in server. And the changelist value in server should not include or be equal to the one in client, which means the changelist value in client 
                    // should be different with server, just like the second example in section 4.6.2.
                    sampleForPidTagPredecessorChangeList = new byte[25];
                    index = 0;
                    length = 0;
                    length = (short)(GidLength + 1);
                    Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                    index += sizeof(short);
                    sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + currentChangeNumber.Length); // 16
                    index += 1;
                    Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                    index += this.localReplicaGuid.ToByteArray().Length;
                    Array.Copy(currentChangeNumber, 0, sampleForPidTagPredecessorChangeList, index, currentChangeNumber.Length);
                    this.lastConflictInfo.PCLB = sampleForPidTagPredecessorChangeList;

                    // The parameter change is used to save the changelist value in client.
                    byte[] change = new byte[48];
                    Array.Copy(BitConverter.GetBytes(46), 0, change, 0, sizeof(short));
                    Array.Copy(sampleForPidTagPredecessorChangeList, 2, change, 2, 23);
                    Array.Copy(BitConverter.GetBytes(22), 0, change, 25, sizeof(short));

                    // Modify the value of PidTagPredecessorChangeList property in client via adding one in GUID field and adding one in GLOBCNT field.
                    byte[] namespaceGuidPartValue = new byte[8];
                    Array.Copy(sampleForPidTagPredecessorChangeList, 10, namespaceGuidPartValue, 2, namespaceGuidPartValue.Length - 2);
                    Array.Reverse(namespaceGuidPartValue);
                    long namespaceGuidPartValueInt = BitConverter.ToInt64(namespaceGuidPartValue, 0);
                    namespaceGuidPartValueInt += 1;
                    namespaceGuidPartValue = BitConverter.GetBytes(namespaceGuidPartValueInt);
                    Array.Reverse(namespaceGuidPartValue);
                    Array.Copy(namespaceGuidPartValue, 2, sampleForPidTagPredecessorChangeList, 10, namespaceGuidPartValue.Length - 2);
                    Array.Copy(sampleForPidTagPredecessorChangeList, 3, change, 26, 22);
                    change[47] += 1;

                    this.lastConflictInfo.PCLA = change;
                    this.lastConflictInfo.PCLXFromMath = change;
                    this.lastConflictInfo.DetectedResult = conflictType;

                    propertyValue.Value = change;
                    hierarchyValues[4] = propertyValue;
                }
            }

            // Construct PidTagDisplayName value.
            propertyValue = new TaggedPropertyValue
            {
                PropertyTag = new PropertyTag(0x3001, 0x001F)
            };

            byte[] sampleForPidTagDisplayName = new byte[Encoding.Unicode.GetByteCount("TestFolder" + "\0")];
            Array.Copy(
                Encoding.Unicode.GetBytes("TestFolder" + "\0"),
                0,
                sampleForPidTagDisplayName,
                0,
                Encoding.Unicode.GetByteCount("TestFolder" + "\0"));
            propertyValue.Value = sampleForPidTagDisplayName;
            hierarchyValues[5] = propertyValue;

            return hierarchyValues;
        }

        /// <summary>
        /// Create values for SynchronizationImportMessageChange
        /// </summary>
        /// <param name="messageIndex">Index of Message ID.</param>
        /// <returns>Message property value of message change.</returns>
        private TaggedPropertyValue[] CreateSamplePropertyValues(int messageIndex)
        {
            byte[] changeNumber = new byte[6];
            short length = 0;
            int index = 0;

            // Send the RopGetLocalReplicaIds request to reserve a range of IDs to be used by a local replica.
            TaggedPropertyValue[] propertyValues = new TaggedPropertyValue[4];
            TaggedPropertyValue propertyValue;
            byte[] currentLocalGloabalCount = this.localId;

            // Set the message property that the message has been already existed.
            if (this.objectIdContainer.ContainsKey(messageIndex))
            {
                propertyValue = new TaggedPropertyValue
                {
                    // Construct PidTagSourceKey value.
                    PropertyTag = new PropertyTag(0x65E0, 0x0102)
                };

                byte[] sample = this.propertyValuesSpecific["PidTagSourceKey"];
                propertyValue.Value = sample;
                propertyValues[0] = propertyValue;

                // Construct PidTagLastModificationTime value.
                propertyValue = new TaggedPropertyValue();
                byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                propertyValue.PropertyTag = new PropertyTag(0x3008, 0x0040);

                propertyValue.Value = sampleForPidTagLastModificationTime;
                propertyValues[1] = propertyValue;

                // Construct PidTagChangeKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E2, 0x0102)
                };

                byte[] sampleForPidTagChangeKey = this.propertyValuesSpecific["PidTagChangeKey"];

                // Modify change key
                byte[] temp = new byte[4];
                Array.Copy(sampleForPidTagChangeKey, sampleForPidTagChangeKey.Length - 4, temp, 0, 4);
                Array.Reverse(temp);
                int add = BitConverter.ToInt32(temp, 0);

                // Add change key to newer version.
                add++;
                temp = BitConverter.GetBytes(add);
                Array.Reverse(temp);
                Array.Copy(temp, 0, sampleForPidTagChangeKey, sampleForPidTagChangeKey.Length - 4, 4);
                propertyValue.Value = sampleForPidTagChangeKey;
                Array.Copy(propertyValue.Value, 2, this.importPidTagChangeKeyValue, 0, propertyValue.Value.Length - 2);
                propertyValues[2] = propertyValue;

                // Construct PidTagPredecessorChangeList value.
                byte[] tempSampleForPidTagChangeKey = new byte[sampleForPidTagChangeKey.Length - 1];
                tempSampleForPidTagChangeKey[0] = sampleForPidTagChangeKey[0];
                Array.Copy(sampleForPidTagChangeKey, 2, tempSampleForPidTagChangeKey, 1, sampleForPidTagChangeKey.Length - 2);
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E3, 0x0102)
                };

                byte[] sampleForPidTagPredecessorChangeList = new byte[this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length + tempSampleForPidTagChangeKey.Length];
                Array.Copy(this.propertyValuesSpecific["PidTagPredecessorChangeList"], 0, sampleForPidTagPredecessorChangeList, 0, this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length);
                Array.Copy(tempSampleForPidTagChangeKey, 0, sampleForPidTagPredecessorChangeList, this.propertyValuesSpecific["PidTagPredecessorChangeList"].Length, tempSampleForPidTagChangeKey.Length);

                // Modify length.
                byte[] tempLength = new byte[2];
                Array.Copy(sampleForPidTagPredecessorChangeList, 0, tempLength, 0, 2);
                short lengthofData = BitConverter.ToInt16(tempLength, 0);
                lengthofData += (short)tempSampleForPidTagChangeKey.Length;
                tempLength = BitConverter.GetBytes(lengthofData);
                Array.Copy(tempLength, 0, sampleForPidTagPredecessorChangeList, 0, 2);
                propertyValue.Value = sampleForPidTagPredecessorChangeList;
                propertyValues[3] = propertyValue;
            }
            else if (messageIndex == 0)
            {
                // Construct PidTagSourceKey value.
                byte[] sub = new byte[4];
                Array.Copy(currentLocalGloabalCount, 2, sub, 0, currentLocalGloabalCount.Length - 2);
                Array.Reverse(sub);
                int totalNum = BitConverter.ToInt32(sub, 0);

                // Generate a new local id from id set got from server.
                totalNum += this.localIdOffSet;
                sub = BitConverter.GetBytes(totalNum);
                Array.Reverse(sub);
                Array.Copy(sub, 0, currentLocalGloabalCount, 2, currentLocalGloabalCount.Length - 2);
                this.localIdOffSet++;

                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E0, 0x0102)
                };

                byte[] sample = new byte[24];

                // The combination of first two bytes (0x0016) indicates the length of value field.
                length = (short)(this.serverReplicaGuid.ToByteArray().Length + currentLocalGloabalCount.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sample, index, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, sample, index, this.serverReplicaGuid.ToByteArray().Length);
                index += this.serverReplicaGuid.ToByteArray().Length;
                Array.Copy(currentLocalGloabalCount, 0, sample, index, currentLocalGloabalCount.Length);
                Array.Copy(currentLocalGloabalCount, 0, changeNumber, 0, currentLocalGloabalCount.Length);
                propertyValue.Value = sample;
                propertyValues[0] = propertyValue;
                length = 0;
                index = 0;

                // Construct PidTagLastModificationTime value.
                propertyValue = new TaggedPropertyValue();
                byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                propertyValue.PropertyTag = new PropertyTag(0x3008, 0x0040);
                propertyValue.Value = sampleForPidTagLastModificationTime;
                propertyValues[1] = propertyValue;

                // Construct PidTagChangeKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E2, 0x0102)
                };
                byte[] sampleForPidTagChangeKey = new byte[24];

                // The combination of first two bytes (0x0014) indicates the length of value field.
                length = (short)(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagChangeKey, 0, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagChangeKey, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagChangeKey, index, changeNumber.Length);
                propertyValue.Value = sampleForPidTagChangeKey;
                propertyValues[2] = propertyValue;
                Array.Copy(propertyValue.Value, 2, this.importPidTagChangeKeyValue, 0, propertyValue.Value.Length - 2);
                length = 0;
                index = 0;

                if (this.propertyValuesSpecific.ContainsKey("LastImportPidTagChangekey"))
                {
                    this.propertyValuesSpecific["LastImportPidTagChangekey"] = sampleForPidTagChangeKey;
                }
                else
                {
                    this.propertyValuesSpecific.Add("LastImportPidTagChangekey", sampleForPidTagChangeKey);
                }

                // Construct PidTagPredecessorChangeList value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E3, 0x0102)
                };

                byte[] sampleForPidTagPredecessorChangeList = new byte[25];
                length = (short)(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length + 1);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                index += sizeof(short);
                sampleForPidTagPredecessorChangeList[index++] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length);
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagPredecessorChangeList, index, changeNumber.Length);
                propertyValue.Value = sampleForPidTagPredecessorChangeList;
                propertyValues[3] = propertyValue;
            }
            else if (messageIndex == -1)
            {
                // Construct PidTagSourceKey value.
                byte[] sub = new byte[4];
                Array.Copy(currentLocalGloabalCount, 2, sub, 0, currentLocalGloabalCount.Length - 2);
                Array.Reverse(sub);
                int totalNum = BitConverter.ToInt32(sub, 0);

                // Generate a new local id from id set got from server.
                totalNum += this.localIdOffSet;
                sub = BitConverter.GetBytes(totalNum);
                Array.Reverse(sub);
                Array.Copy(sub, 0, currentLocalGloabalCount, 2, currentLocalGloabalCount.Length - 2);
                this.localIdOffSet++;

                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E0, 0x0102)
                };

                byte[] sample = new byte[24];

                // The combination of first two bytes (0x0016) indicates the length of value field.
                length = (short)(this.serverReplicaGuid.ToByteArray().Length + currentLocalGloabalCount.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sample, index, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, sample, index, this.serverReplicaGuid.ToByteArray().Length);
                index += this.serverReplicaGuid.ToByteArray().Length;
                Array.Copy(currentLocalGloabalCount, 0, sample, index, currentLocalGloabalCount.Length);
                Array.Copy(currentLocalGloabalCount, 0, changeNumber, 0, currentLocalGloabalCount.Length);
                propertyValue.Value = sample;
                propertyValues[0] = propertyValue;
                length = 0;
                index = 0;

                // Construct PidTagLastModificationTime value.
                propertyValue = new TaggedPropertyValue();
                byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                propertyValue.PropertyTag = new PropertyTag(0x3008, 0x0040);
                propertyValue.Value = sampleForPidTagLastModificationTime;
                propertyValues[1] = propertyValue;

                // Construct PidTagChangeKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E2, 0x0102)
                };
                byte[] sampleForPidTagChangeKey = new byte[24];

                // The combination of first two bytes (0x0014) indicates the length of value field.
                length = (short)(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagChangeKey, 0, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagChangeKey, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagChangeKey, index, changeNumber.Length);

                // The parameter localIDs is used to save the LocalId value of the PidTagChangeKey value. 
                byte[] localIDs = new byte[8];
                Array.Copy(sampleForPidTagChangeKey, 18, localIDs, 2, localIDs.Length - 2);
                Array.Reverse(localIDs);
                long localIDsInt = BitConverter.ToInt64(localIDs, 0);
                localIDsInt += 1;
                localIDs = BitConverter.GetBytes(localIDsInt);
                Array.Reverse(localIDs);
                Array.Copy(localIDs, 2, sampleForPidTagChangeKey, 18, localIDs.Length - 2);

                propertyValue.Value = sampleForPidTagChangeKey;
                propertyValues[2] = propertyValue;
                Array.Copy(propertyValue.Value, 2, this.importPidTagChangeKeyValue, 0, propertyValue.Value.Length - 2);
                length = 0;
                index = 0;

                // Construct PidTagPredecessorChangeList value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E3, 0x0102)
                };

                byte[] sampleForPidTagPredecessorChangeList = new byte[25];
                index = 0;
                length = 0;
                length = (short)(GidLength + 1);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                index += sizeof(short);
                sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length); // 16
                index += 1;
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagPredecessorChangeList, index, changeNumber.Length);

                this.lastConflictInfo.PCLB = sampleForPidTagPredecessorChangeList;
                sampleForPidTagPredecessorChangeList[24] += 1;
                this.lastConflictInfo.PCLA = sampleForPidTagPredecessorChangeList;
                this.lastConflictInfo.PCLXFromMath = sampleForPidTagPredecessorChangeList;
                this.lastConflictInfo.DetectedResult = ConflictTypes.AINCLUDEB;

                propertyValue.Value = sampleForPidTagPredecessorChangeList;
                propertyValues[3] = propertyValue;
            }
            else if (messageIndex == -2)
            {
                byte[] sub = new byte[4];
                Array.Copy(currentLocalGloabalCount, 2, sub, 0, currentLocalGloabalCount.Length - 2);
                Array.Reverse(sub);
                int totalNum = BitConverter.ToInt32(sub, 0);

                // Generate a new local id from id set got from server.
                totalNum += this.localIdOffSet;
                sub = BitConverter.GetBytes(totalNum);
                Array.Reverse(sub);
                Array.Copy(sub, 0, currentLocalGloabalCount, 2, currentLocalGloabalCount.Length - 2);
                this.localIdOffSet++;

                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E0, 0x0102)
                };

                byte[] sample = new byte[24];

                // The combination of first two bytes (0x0016) indicates the length of value field.
                length = (short)(this.serverReplicaGuid.ToByteArray().Length + currentLocalGloabalCount.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sample, index, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.serverReplicaGuid.ToByteArray(), 0, sample, index, this.serverReplicaGuid.ToByteArray().Length);
                index += this.serverReplicaGuid.ToByteArray().Length;
                Array.Copy(currentLocalGloabalCount, 0, sample, index, currentLocalGloabalCount.Length);
                Array.Copy(currentLocalGloabalCount, 0, changeNumber, 0, currentLocalGloabalCount.Length);
                propertyValue.Value = sample;
                propertyValues[0] = propertyValue;
                length = 0;
                index = 0;

                // Construct PidTagLastModificationTime value.
                propertyValue = new TaggedPropertyValue();
                byte[] sampleForPidTagLastModificationTime = BitConverter.GetBytes(DateTime.Now.ToFileTimeUtc());
                propertyValue.PropertyTag = new PropertyTag(0x3008, 0x0040);
                propertyValue.Value = sampleForPidTagLastModificationTime;
                propertyValues[1] = propertyValue;

                // Construct PidTagChangeKey value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E2, 0x0102)
                };
                byte[] sampleForPidTagChangeKey = new byte[24];

                // The combination of first two bytes (0x0014) indicates the length of value field.
                length = (short)(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagChangeKey, 0, sizeof(short));
                index += sizeof(short);
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagChangeKey, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagChangeKey, index, changeNumber.Length);

                // The parameter localIDs is used to save the LocalId value of the PidTagChangeKey value. 
                byte[] localIDs = new byte[8];
                Array.Copy(sampleForPidTagChangeKey, 18, localIDs, 2, localIDs.Length - 2);
                Array.Reverse(localIDs);
                long localIDsInt = BitConverter.ToInt64(localIDs, 0);
                localIDsInt += 1;
                localIDs = BitConverter.GetBytes(localIDsInt);
                Array.Reverse(localIDs);
                Array.Copy(localIDs, 2, sampleForPidTagChangeKey, 18, localIDs.Length - 2);
                propertyValue.Value = sampleForPidTagChangeKey;
                propertyValues[2] = propertyValue;
                Array.Copy(propertyValue.Value, 2, this.importPidTagChangeKeyValue, 0, propertyValue.Value.Length - 2);
                length = 0;
                index = 0;

                // Construct PidTagPredecessorChangeList value.
                propertyValue = new TaggedPropertyValue
                {
                    PropertyTag = new PropertyTag(0x65E3, 0x0102)
                };

                byte[] sampleForPidTagPredecessorChangeList = new byte[25];
                index = 0;
                length = 0;
                length = (short)(GidLength + 1);
                Array.Copy(BitConverter.GetBytes(length), 0, sampleForPidTagPredecessorChangeList, 0, sizeof(short));
                index += sizeof(short);
                sampleForPidTagPredecessorChangeList[2] = Convert.ToByte(this.localReplicaGuid.ToByteArray().Length + changeNumber.Length); // 16
                index += 1;
                Array.Copy(this.localReplicaGuid.ToByteArray(), 0, sampleForPidTagPredecessorChangeList, index, this.localReplicaGuid.ToByteArray().Length);
                index += this.localReplicaGuid.ToByteArray().Length;
                Array.Copy(changeNumber, 0, sampleForPidTagPredecessorChangeList, index, changeNumber.Length);

                this.lastConflictInfo.PCLB = sampleForPidTagPredecessorChangeList;

                byte[] change = new byte[48];
                Array.Copy(BitConverter.GetBytes(46), 0, change, 0, sizeof(short));
                Array.Copy(sampleForPidTagPredecessorChangeList, 2, change, 2, 23);
                Array.Copy(BitConverter.GetBytes(22), 0, change, 25, sizeof(short));
                Array.Copy(Guid.NewGuid().ToByteArray(), 0, change, 26, 16);

                this.lastConflictInfo.PCLA = change;
                this.lastConflictInfo.PCLXFromMath = change;

                propertyValue.Value = change;
                propertyValues[3] = propertyValue;
            }

            return propertyValues;
        }

        /// <summary>
        /// Get required buffer from total buffer.
        /// </summary>
        /// <param name="buf">Total buffer needs to be processed.</param>
        /// <returns>Processed buffer.</returns>
        private byte[] ProcessFXSourceGetBuffer(byte[] buf)
        {
            byte[] bufferMarker = new byte[4];
            int j = 0;
            for (j = 0; j < buf.Length; j++)
            {
                if (j <= buf.Length - 4)
                {
                    Array.Copy(buf, j, bufferMarker, 0, 4);
                    uint marker = BitConverter.ToUInt32(bufferMarker, 0);
                    if ((Markers)marker == Markers.PidTagStartTopFld)
                    {
                        break;
                    }
                }
            }

            if (j == buf.Length)
            {
                j = 0;
                if (buf[3] == 0x40 && buf[2] == 0x08 && buf[1] == 0x84 && buf[0] == 0xe4)
                {
                    j = buf[4] + 8;
                }
            }

            byte[] processedBuf = new byte[buf.Length - j];
            Array.Copy(buf, j, processedBuf, 0, processedBuf.Length);
            return processedBuf;
        }

        /// <summary>
        /// Revert public folder by deleting its subfolder and messages.
        /// </summary>
        /// <param name="logonFlags"> Indicate the logon flag.</param>
        private void RevertPublicFolder(LogonFlags logonFlags)
        {
            int logonHandleIndex, inboxFolderIndex;

            // Reconnect the server
            bool resultDisconnect = this.oxcropsClient[1].Disconnect();
            Site.Assert.IsTrue(resultDisconnect, "Disconnecting server should be successful.");
            this.Connect(1, Microsoft.Protocols.TestSuites.Common.ConnectionType.PublicFolderServer);
            this.Logon(1, logonFlags, out logonHandleIndex, out inboxFolderIndex);
            List<uint> soh;
            List<uint> sohGetHierarchyFirst;
            RopOpenFolderResponse res = this.OpenFolder(1, (uint)this.logonHandleOut, this.rootFolderId, out soh);
            uint rootPublicFolderHandle = soh[res.OutputHandleIndex];
            RopGetHierarchyTableResponse getHirerarchyTableResFirst = this.GetHierarchyTable(1, rootPublicFolderHandle, TableFlags.None, out sohGetHierarchyFirst);
            this.DeleteFolder(1, rootPublicFolderHandle, this.publicFolderId);
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
            while (retryCount > 0)
            {
                List<uint> sohGetHierarchySecond;
                RopGetHierarchyTableResponse getHirerarchyTableResSecond = this.GetHierarchyTable(1, rootPublicFolderHandle, TableFlags.None, out sohGetHierarchySecond);
                if (getHirerarchyTableResSecond.RowCount + 1 == getHirerarchyTableResFirst.RowCount)
                {
                    break;
                }

                System.Threading.Thread.Sleep(waitTime);
                retryCount--;
            }

            if (retryCount == 0)
            {
                Site.Assert.Fail("The public folder still exists after {0} seconds.", waitTime * retryCount / 1000);
            }

            if (logonFlags == LogonFlags.PublicFolder)
            {
                RopResult result = RopResult.InvalidParameter;
                int folderIdIndex, folderHandleIndex;
                result = this.CreateFolder(1, rootPublicFolderHandle, Common.GetConfigurationPropertyValue("PublicFolderName", this.Site), out folderIdIndex, out folderHandleIndex);
                Site.Assert.AreEqual<RopResult>(RopResult.Success, result, "Fail to create public folder, the error code is {0}", result.ToString());
            }
        }

        /// <summary>
        /// Get the data from stack.
        /// </summary>
        /// <param name="maxIndex">The max index in the command list.</param>
        /// <param name="commandList">The data source of command list.</param>
        /// <param name="comByteStack">The initial result returned.</param>
        /// <returns>The final data get from the data source based on the initial data.</returns>
        private CommonByteStack GetCommonByteStack(int maxIndex, List<Command> commandList, CommonByteStack comByteStack)
        {
            CommonByteStack commonByteStack = comByteStack;

            // Check the commands before current one in DeserializedCommandList,
            for (int j = 0; j < maxIndex; j++)
            {
                if (commandList[j] is PushCommand)
                {
                    // if  encounter PushCommand, push the element into comByteStack
                    commonByteStack.Push((commandList[j] as PushCommand).CommonBytes);

                    // When a Push command places a sixth byte onto the common byte stack, 
                    // it tells the decoder the next GLOBCNT pair has all six bytes in common
                    if (commonByteStack.Bytes == 6)
                    {
                        commonByteStack.Pop();
                    }
                }
                else if (commandList[j] is PopCommand)
                {
                    // If encounter PopCommand, pop the elements out of comByteStac
                    commonByteStack.Pop();
                }
            }

            return commonByteStack;
        }
        #endregion
    }
}