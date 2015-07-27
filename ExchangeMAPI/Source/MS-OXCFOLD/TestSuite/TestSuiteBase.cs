//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The base test suite class defines common initialization method and cleanup method for all 5 scenarios.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : Microsoft.Protocols.TestTools.TestClassBase
    {
        /// <summary>
        /// An instance of protocol adapter.
        /// </summary>
        private IMS_OXCFOLDAdapter adapter;

        /// <summary>
        /// Indicates whether root folder created successfully in test case.
        /// </summary>
        private bool isRootFolderCreated = false;

        /// <summary>
        /// Handle for a Mailbox logon.
        /// </summary>
        private uint logonHandle;

        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

        /// <summary>
        /// The handle of the inbox folder.
        /// </summary>
        private uint inboxFolderHandle;

        /// <summary>
        /// The handle of the root folder which is used by MS-OXCFOLD test cases.
        /// This folder is created for MS-OXCFOLD test suite under which all other folders
        /// and messages for test will be created.
        /// </summary>
        private uint rootFolderHandle;

        /// <summary>
        /// The ID of the root folder which is used by MS-OXCFOLD test cases.
        /// This folder is created for MS-OXCFOLD test suite under which all other folders
        /// and messages for test will be created.
        /// </summary>
        private ulong rootFolderId;

        /// <summary>
        /// 13 folder IDs get in the RopLogon rop response, according to MS-OXCSTOR 2.2.1.1.3 "RopLogon ROP Success Response Buffer for Private Mailbox"
        /// </summary>
        private ulong[] defaultFolderIds;

        /// <summary>
        /// The time internal (millisecond) that is used by the test suite to find the message item after it sends out.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// The retry count that is used by the test suite to find the message item after it sends out.
        /// </summary>
        private int retryCount;

        /// <summary>
        /// The name of the root folder which is used by MS-OXCFOLD test cases.
        /// </summary>
        private string rootFolder;

        /// <summary>
        /// A Boolean indicates the cases whether need call cleanup.
        /// </summary>
        private bool needCleanup = true;
        #region Properties

        /// <summary>
        /// Gets 13 folder IDs get in the RopLogon rop response.
        /// </summary>
        public ulong[] DefaultFolderIds
        {
            get
            {
                return this.defaultFolderIds;
            }
        }

        /// <summary>
        /// Gets the logon handle for a MailBox.
        /// </summary>
        public uint LogonHandle
        {
            get
            {
                return this.logonHandle;
            }
        }

        /// <summary>
        /// Gets or sets the name of the root folder which is used by MS-OXCFOLD test cases.
        /// </summary>
        protected string RootFolder
        {
            get
            {
                return this.rootFolder;
            }

            set
            {
                this.rootFolder = value;
            }
        }

        /// <summary>
        /// Gets or sets the reference to the protocol adapter instance.
        /// </summary>
        protected IMS_OXCFOLDAdapter Adapter
        {
            get
            {
                return this.adapter;
            }

            set
            {
                this.adapter = value;
            }
        }

        /// <summary>
        /// Gets the handle of the root folder.
        /// </summary>
        protected uint RootFolderHandle
        {
            get
            {
                return this.rootFolderHandle;
            }
        }

        /// <summary>
        /// Gets the ID of the root folder.
        /// </summary>
        protected ulong RootFolderId
        {
            get
            {
                return this.rootFolderId;
            }
        }

        /// <summary>
        /// Gets the value of the waitTime.
        /// </summary>
        protected int WaitTime
        {
            get
            {
                return this.waitTime;
            }
        }

        /// <summary>
        /// Gets the value of the retryCount.
        /// </summary>
        protected int RetryCount
        {
            get
            {
                return this.retryCount;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the cases whether need call cleanup.
        /// </summary>
        protected bool NeedCleanup
        {
            get { return this.needCleanup; }
            set { this.needCleanup = value; }
        }
        #endregion

        #region Test Case Initialization
        /// <summary>
        /// Test initialize. Overrides the method TestInitialize defined in base class.
        /// </summary>
        protected override void TestInitialize()
        {
            this.adapter = Site.GetAdapter<IMS_OXCFOLDAdapter>();
            this.waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            this.retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            this.rootFolder = Common.GenerateResourceName(this.Site, Constants.RootFolder) + Constants.StringNullTerminated;
            this.needCleanup = true;
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            bool transportIsMAPI = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http";
            if (!transportIsMAPI || Common.IsRequirementEnabled(1340, this.Site))
            {
                if (!this.needCleanup)
                {
                    return;
                }

                #region Hard Delete Messages and Subfolders of Inbox
                if (this.isRootFolderCreated)
                {
                    RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest();
                    RopDeleteFolderResponse deleteFolderResponse;
                    deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
                    deleteFolderRequest.LogonId = Constants.CommonLogonId;
                    deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;

                    // Set the DeleteFolderFlags to indicate hard delete the common generic folder,
                    // including all folders and messages under the folder.
                    deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders
                        | (byte)DeleteFolderFlags.DelMessages
                        | (byte)DeleteFolderFlags.DeleteHardDelete;
                    deleteFolderRequest.FolderId = this.RootFolderId;

                    int count = 0;
                    bool rootFolderCleanUpSuccess = false;
                    do
                    {
                        deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, this.inboxFolderHandle, ref this.responseHandles);
                        if (deleteFolderResponse.ReturnValue == Constants.SuccessCode && deleteFolderResponse.PartialCompletion == 0)
                        {
                            rootFolderCleanUpSuccess = true;
                        }
                        else
                        {
                            Thread.Sleep(this.WaitTime);
                        }

                        if (count > this.RetryCount)
                        {
                            break;
                        }

                        count++;
                    }
                    while (!rootFolderCleanUpSuccess);

                    Site.Assert.IsTrue(rootFolderCleanUpSuccess, "Test case cleanup should succeed.");
                }               
                #endregion

                #region  RopRelease
                RopReleaseRequest releaseRequest = new RopReleaseRequest();
                object ropResponse = null;
                releaseRequest.RopId = (byte)RopId.RopRelease;
                releaseRequest.LogonId = Constants.CommonLogonId;
                releaseRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
                this.Adapter.DoRopCall(releaseRequest, this.LogonHandle, ref ropResponse, ref this.responseHandles);
                #endregion

                this.logonHandle = 0;
                ropResponse = null;
                this.responseHandles = null;
                this.Adapter.DoDisconnect();
            }
        }

        #endregion

        #region Common Private Type(s) and Function(s)
        
        /// <summary>
        /// Check whether support MS-OXCMAPIHTTP or ncacn_ip_tcp transport.
        /// </summary>
        protected void CheckWhetherSupportTransport()
        {
            if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http") && (!Common.IsRequirementEnabled(1340, this.Site)))
            {
                this.needCleanup = false;
                Site.Assert.Inconclusive("This case can not run because MAPIHTTP is not enabled to run. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
            else if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "ncacn_ip_tcp") && (!Common.IsRequirementEnabled(99999, this.Site)))
            {
                this.needCleanup = false;
                Site.Assert.Inconclusive("This case can not run because ncacn_ip_tcp is not enabled to run. ncacn_ip_tcp is not supported from Exchange Server 2013.");
            }
        }

        /// <summary>
        /// Initialize a generic folder under inbox as a root folder for test.
        /// </summary>
        protected void GenericFolderInitialization()
        {
            RopLogonResponse logonResponse = this.Logon(LogonFlags.Private, out this.logonHandle);

            #region Open the inbox folder.
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderId = logonResponse.FolderIds[Constants.InboxIndex],
                OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted
            };

            // Use the logon object as input handle here.
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);
            this.inboxFolderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];
            #endregion

            #region Create a generic folder for test.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest();
            RopCreateFolderResponse createFolderResponse = new RopCreateFolderResponse();
            createFolderRequest.RopId = (byte)RopId.RopCreateFolder;
            createFolderRequest.LogonId = Constants.CommonLogonId;
            createFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createFolderRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createFolderRequest.FolderType = 0x01;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0x01;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(this.RootFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(this.RootFolder);
            createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.inboxFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating Folder should succeed.");
            this.rootFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            this.rootFolderId = createFolderResponse.FolderId;
            this.isRootFolderCreated = true;
            #endregion
        }

        /// <summary>
        /// Logon Mail server.
        /// </summary>
        /// <param name="logonType">Logon Type.</param>
        /// <param name="objHandle">Server response handle.</param>
        /// <returns>Logon Response. </returns>
        protected RopLogonResponse Logon(LogonFlags logonType, out uint objHandle)
        {
            uint openFlags = (uint)OpenFlags.None;
            if (LogonFlags.PublicFolder == logonType)
            {
                openFlags = (uint)OpenFlags.UsePerMDBReplipMapping | (uint)OpenFlags.IgnoreHomeMDB | (uint)OpenFlags.Public;
            }
            else if (LogonFlags.Private == logonType)
            {
                openFlags = (uint)OpenFlags.UsePerMDBReplipMapping | (uint)OpenFlags.UseAdminPrivilege;
            }

            return this.Logon(logonType, out objHandle, openFlags);
        }

        /// <summary>
        /// Logon Mail server.
        /// </summary>
        /// <param name="logonType">Logon Type.</param>
        /// <param name="objHandle">Server response handle.</param>
        /// <param name="openFlags">The possible values are specified in [MS-OXCSTOR]. This structure contains more flags that control the behavior of the logon.</param>
        /// <returns>Logon Response.</returns>
        protected RopLogonResponse Logon(LogonFlags logonType, out uint objHandle, uint openFlags)
        {
            RopLogonRequest logonRequest = new RopLogonRequest();
            object ropResponse = null;

            string userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site) + Constants.StringNullTerminated;

            logonRequest.RopId = (byte)RopId.RopLogon;
            logonRequest.LogonId = Constants.CommonLogonId;
            logonRequest.OutputHandleIndex = 0x0;
            logonRequest.StoreState = 0;
            logonRequest.LogonFlags = (byte)logonType;
            logonRequest.OpenFlags = openFlags;

            if (LogonFlags.PublicFolder == logonType)
            {
                logonRequest.EssdnSize = 0;
                logonRequest.Essdn = null;
            }
            else if (LogonFlags.Private == logonType)
            {
                logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);
                logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);
            }

            this.Adapter.DoRopCall(logonRequest, this.LogonHandle, ref ropResponse, ref this.responseHandles);
            RopLogonResponse logonResponse = (RopLogonResponse)ropResponse;
            objHandle = this.responseHandles[0][logonResponse.OutputHandleIndex];
            this.defaultFolderIds = logonResponse.FolderIds;
            return logonResponse;
        }

        /// <summary>
        /// Opens an existing folder.
        /// </summary>
        /// <param name="inputHandle">The input Server object for this operation is a Logon object or a Folder object that represents the object to be opened.</param>
        /// <param name="folderId">The folder ID specifies the folder to be opened.</param>
        /// <param name="folderHandle">A Folder object that represents the folder that was opened.</param>
        /// <returns>The RopOpenFolder ROP response buffer.</returns>
        protected RopOpenFolderResponse OpenFolder(uint inputHandle, ulong folderId, ref uint folderHandle)
        {
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                OpenModeFlags = (byte)FolderOpenModeFlags.None,
                FolderId = folderId
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, inputHandle, ref this.responseHandles);
            folderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];
            return openFolderResponse;
        }

        /// <summary>
        /// Creates a new Folder.
        /// </summary>
        /// <param name="parentFolderHandle">The parent folder under which the new folder will be created.</param>
        /// <param name="folderName">Folder Name which will be created.</param>
        /// <param name="folderId">Folder Id which will be returned by server.</param>
        /// <param name="folderHandle">Folder Handle which will be returned by server.</param>
        /// <returns>The RopCreateFolder ROP response buffer.</returns>
        protected RopCreateFolderResponse CreateFolder(uint parentFolderHandle, string folderName, ref ulong folderId, ref uint folderHandle)
        {
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(folderName),
                Comment = Encoding.ASCII.GetBytes(folderName)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, parentFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "Creating Folder should succeed.");
            folderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            folderId = createFolderResponse.FolderId;
            return createFolderResponse;
        }

        /// <summary>
        /// Create and save message.
        /// </summary>
        /// <param name="folderHandle">Folder Handle </param>
        /// <param name="folderId">Folder Id  which messages will be create in.</param>
        /// <param name="messageId"> Message Id which will be returned by server.</param>
        /// <param name="messageHandle">Message Handle which will be returned by server.</param>
        protected void CreateSaveMessage(uint folderHandle, ulong folderId, ref ulong messageId, ref uint messageHandle)
        {
            this.CreateSaveMessage(folderHandle, folderId, 0, ref messageId, ref messageHandle);
        }

        /// <summary>
        /// Create and save message.
        /// </summary>
        /// <param name="folderHandle">Folder Handle </param>
        /// <param name="folderId">Folder Id  which messages will be create in.</param>
        /// <param name="associatedFlag">Specifies whether the message is a Folder Associated Information message.</param>
        /// <param name="messageId"> Message Id which will be returned by server.</param>
        /// <param name="messageHandle">Message Handle which will be returned by server.</param>
        protected void CreateSaveMessage(uint folderHandle, ulong folderId, byte associatedFlag, ref ulong messageId, ref uint messageHandle)
        {
            // Create a Message.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse = new RopCreateMessageResponse();
            object ropResponse = null;
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = Constants.CommonLogonId;
            createMessageRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            createMessageRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            createMessageRequest.CodePageId = 0x0FFF;
            createMessageRequest.FolderId = folderId;
            createMessageRequest.AssociatedFlag = associatedFlag;
            this.Adapter.DoRopCall(createMessageRequest, folderHandle, ref ropResponse, ref this.responseHandles);

            createMessageResponse = (RopCreateMessageResponse)ropResponse;
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "Creating Message should succeed");
            messageHandle = this.responseHandles[0][createMessageResponse.OutputHandleIndex];

            // Save a Message.
            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest();
            RopSaveChangesMessageResponse saveChangesMessageResponse = new RopSaveChangesMessageResponse();
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = Constants.CommonLogonId;
            saveChangesMessageRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            saveChangesMessageRequest.ResponseHandleIndex = 0x01;
            saveChangesMessageRequest.SaveFlags = 0x0C;
            this.Adapter.DoRopCall(saveChangesMessageRequest, messageHandle, ref ropResponse, ref this.responseHandles);

            saveChangesMessageResponse = (RopSaveChangesMessageResponse)ropResponse;
            Site.Assert.AreEqual<uint>(
                0,
                createMessageResponse.ReturnValue,
                "Save Messages Success.");
            messageId = saveChangesMessageResponse.MessageId;
        }

        /// <summary>
        /// Create sample RuleData array to modify the rules associated with a folder.
        /// </summary>
        /// <returns>Return RuleData array to be used in RopModifyRules request.</returns>
        protected RuleData[] CreateSampleRuleDataArrayForAdd()
        {
            int propertyValuesCount = 4;
            PropertyValue[] propertyValues = new PropertyValue[propertyValuesCount];

            for (int i = 0; i < propertyValuesCount; i++)
            {
                propertyValues[i] = new PropertyValue();
            }

            // As specified in section 2.2.1.3.2 in [MS-OXORULE],
            // when adding a PRULE, the client MUST NOT
            // pass in PidTagRuleId, it MUST pass in PidTagRuleCondition,
            // PidTagRuleActions and PidTagRuleProvider.
            TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue();
            PropertyTag tempPropertyTag = new PropertyTag
            {
                PropertyId = 0x6676, PropertyType = 0003
            };
            taggedPropertyValue.PropertyTag = tempPropertyTag;
            byte[] value3 = { 0x00, 0x00, 0x00, 0x0a };
            taggedPropertyValue.Value = value3;
            propertyValues[3].Value = taggedPropertyValue.Serialize();

            // PidTagRuleCondition
            taggedPropertyValue = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = 0x6679;
            tempPropertyTag.PropertyType = 0x00fd;
            taggedPropertyValue.PropertyTag = tempPropertyTag;
            byte[] value =
            {
                0x03, 0x01, 0x00, 0x01, 0x00, 0x1f, 0x00, 0x37, 0x00, 0x1f, 0x00,
                0x37, 0x00, 0x50, 0x00, 0x72, 0x00, 0x6f, 0x00, 0x6a, 0x00, 0x65,
                0x00, 0x63, 0x00, 0x74, 0x00, 0x20, 0x00, 0x58, 0x00, 0x00, 0x00
            };
            taggedPropertyValue.Value = value;
            propertyValues[1].Value = taggedPropertyValue.Serialize();

            // PidTagRuleAction
            taggedPropertyValue = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = 0x6680;
            tempPropertyTag.PropertyType = 0x00fe;
            taggedPropertyValue.PropertyTag = tempPropertyTag;
            byte[] value1 = { 0x01, 0x00, 0x09, 0x00, 0x0B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
            taggedPropertyValue.Value = value1;
            propertyValues[2].Value = taggedPropertyValue.Serialize();

            // PidTagRuleProvider
            taggedPropertyValue = new TaggedPropertyValue();
            tempPropertyTag.PropertyId = 0x6681;
            tempPropertyTag.PropertyType = 0x001f;
            taggedPropertyValue.PropertyTag = tempPropertyTag;
            byte[] value2 = Encoding.Unicode.GetBytes("RuleOrganizer\0");
            taggedPropertyValue.Value = value2;
            propertyValues[0].Value = taggedPropertyValue.Serialize();

            RuleData sampleRuleData = new RuleData
            {
                RuleDataFlags = 0x01,
                PropertyValueCount = (ushort)propertyValues.Length,
                PropertyValues = propertyValues
            };

            RuleData[] sampleRuleDataArray = new RuleData[1];
            sampleRuleDataArray[0] = sampleRuleData;

            return sampleRuleDataArray;
        }

        /// <summary>
        /// Get row count of a specified folder's contents table.
        /// </summary>
        /// <param name="flags">The TableFlags parameter contains a bitmask of flags that control how information is returned in the table on folder.</param>
        /// <param name="inputHandle">The folder object handle.</param>
        /// <returns>Row count of the specified folder's contents table.</returns>
        protected uint GetContentsTable(FolderTableFlags flags, uint inputHandle)
        {
            RopGetContentsTableRequest request = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)flags
            };
            RopGetContentsTableResponse response = this.Adapter.GetContentsTable(request, inputHandle, ref this.responseHandles);
            return response.RowCount;
        }

        /// <summary>
        /// Open a specified message in specified folder.
        /// </summary>
        /// <param name="messageId">The specified message ID.</param>
        /// <param name="folderId">The specified folder ID.</param>
        /// <param name="openModeFlags">OpenModeFlags for RopOpenMessage.</param>
        /// <returns>The response of RopOpenMessage ROP operation.</returns>
        protected RopOpenMessageResponse OpenMessage(ulong messageId, ulong folderId, MessageOpenModeFlags openModeFlags)
        {
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                OpenModeFlags = (byte)FolderOpenModeFlags.None,
                FolderId = folderId
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);
            uint folderHandle = this.responseHandles[0][openFolderResponse.OutputHandleIndex];

            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest();
            object ropResponse = null;
            openMessageRequest.RopId = (byte)RopId.RopOpenMessage;
            openMessageRequest.LogonId = Constants.CommonLogonId;
            openMessageRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            openMessageRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            openMessageRequest.OpenModeFlags = (byte)openModeFlags;
            openMessageRequest.FolderId = folderId;
            openMessageRequest.MessageId = messageId;
            openMessageRequest.CodePageId = 0x0fff;
            this.Adapter.DoRopCall(openMessageRequest, folderHandle, ref ropResponse, ref this.responseHandles);
            return (RopOpenMessageResponse)ropResponse;
        }

        /// <summary>
        /// Verify whether the specified folder was deleted.
        /// </summary>
        /// <param name="folderId">The specified folder ID.</param>
        /// <returns>Return a Boolean value indicates whether the specified folder was deleted.</returns>
        protected bool IsFolderDeleted(ulong folderId)
        {
            bool isFolderDeleted = false;

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                OpenModeFlags = (byte)FolderOpenModeFlags.None,
                FolderId = folderId
            };
            RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);

            if (openFolderResponse.ReturnValue == 0x8004010F)
            {
                isFolderDeleted = true;
            }

            return isFolderDeleted;
        }

        /// <summary>
        /// Verify whether the specified message in specified folder was deleted.
        /// </summary>
        /// <param name="messageId">The specified message ID.</param>
        /// <param name="folderId">The specified folder ID.</param>
        /// <returns>Return a Boolean value indicates whether the specified message was deleted.</returns>
        protected bool IsMessageDeleted(ulong messageId, ulong folderId)
        {
            bool isMessageDeleted = false;

            RopOpenMessageResponse openMessageResponse = this.OpenMessage(messageId, folderId, MessageOpenModeFlags.BestAccess);

            if (openMessageResponse.ReturnValue == 0x8004010F)
            {
                isMessageDeleted = true;
            }

            return isMessageDeleted;
        }

        /// <summary>
        /// Verify whether the specified folder was soft deleted.
        /// </summary>
        /// <param name="folderId">The specified folder ID.</param>
        /// <returns>Return a Boolean value indicates whether the specified folder was soft deleted.</returns>
        protected bool IsFolderSoftDeleted(ulong folderId)
        {
            bool folderDeleted = this.IsFolderDeleted(folderId);
            bool folderSoftDeleted = false;

            if (folderDeleted)
            {
                RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
                {
                    RopId = (byte)RopId.RopOpenFolder,
                    LogonId = Constants.CommonLogonId,
                    InputHandleIndex = Constants.CommonInputHandleIndex,
                    OutputHandleIndex = Constants.CommonOutputHandleIndex,
                    OpenModeFlags = (byte)FolderOpenModeFlags.OpenSoftDeleted,
                    FolderId = folderId
                };
                RopOpenFolderResponse openFolderResponse = this.Adapter.OpenFolder(openFolderRequest, this.LogonHandle, ref this.responseHandles);

                if (openFolderResponse.ReturnValue == 0)
                {
                    folderSoftDeleted = true;
                }
            }

            return folderSoftDeleted;
        }

        /// <summary>
        /// Verify whether the specified folder was hard deleted.
        /// </summary>
        /// <param name="folderId">The specified folder ID.</param>
        /// <returns>Return a Boolean value indicates whether the specified folder was hard deleted.</returns>
        protected bool IsFolderHardDeleted(ulong folderId)
        {
            bool folderDeleted = this.IsFolderDeleted(folderId);
            bool folderSoftDeleted = this.IsFolderSoftDeleted(folderId);

            bool folderHardDeleted = folderDeleted && !folderSoftDeleted;

            return folderHardDeleted;
        }

        /// <summary>
        /// Get the properties' value from the rows of the table.
        /// </summary>
        /// <param name="tableHandle">The table handle.</param>
        /// <param name="rowCount">The amount of the rows.</param>
        /// <param name="properties">The properties need to show.</param>
        /// <returns>The property rows in the specified table object.</returns>
        protected List<PropertyRow> GetTableRowValue(uint tableHandle, ushort rowCount, PropertyTag[] properties)
        {
            #region The client calls RopSetColumns operation to set the property information to show.

            RopSetColumnsRequest setColumnsRequest = new RopSetColumnsRequest();
            object ropResponse = new object();
            setColumnsRequest.RopId = (byte)RopId.RopSetColumns;
            setColumnsRequest.LogonId = Constants.CommonLogonId;
            setColumnsRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            setColumnsRequest.PropertyTagCount = (ushort)properties.Length;
            setColumnsRequest.PropertyTags = properties;
            setColumnsRequest.SetColumnsFlags = (byte)AsynchronousFlags.None;
            this.Adapter.DoRopCall(setColumnsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            #endregion

            #region The client calls RopQueryRows operation to query the folder which have the special properties.

            RopQueryRowsRequest queryRowsRequest = new RopQueryRowsRequest();
            ropResponse = new object();
            queryRowsRequest.RopId = (byte)RopId.RopQueryRows;
            queryRowsRequest.LogonId = Constants.CommonLogonId;
            queryRowsRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            queryRowsRequest.RowCount = (ushort)rowCount;
            queryRowsRequest.QueryRowsFlags = (byte)QueryRowsFlags.Advance;
            queryRowsRequest.ForwardRead = 0x01;
            this.Adapter.DoRopCall(queryRowsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)ropResponse;
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, queryRowsResponse.ReturnValue, "RopQueryRows ROP operation performs successful!");
            #endregion

            List<PropertyRow> propertyRows = null;

            if (queryRowsResponse.RowData != null)
            {
                propertyRows = queryRowsResponse.RowData.PropertyRows;
            }

            return propertyRows;
        }

        /// <summary>
        /// Find a folder ID in the specified folder and with specified display name.
        /// </summary>
        /// <param name="parentFolderId">ID of the parent Folder.</param>
        /// <param name="logonHandle">The logon object handle.</param>
        /// <param name="folderName">The folder display name.</param>
        /// <returns>The folder ID.</returns>
        protected ulong GetSubfolderIDByName(ulong parentFolderId, uint logonHandle, string folderName)
        {
            ulong folderId = 0;
            uint parentFolderHandle = 0;
            this.OpenFolder(logonHandle, parentFolderId, ref parentFolderHandle);

            RopGetHierarchyTableRequest getHierarchyTableRequest = new RopGetHierarchyTableRequest
            {
                RopId = (byte)RopId.RopGetHierarchyTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.Depth
            };
            RopGetHierarchyTableResponse getHierarchyTableResponse = this.Adapter.GetHierarchyTable(getHierarchyTableRequest, parentFolderHandle, ref this.responseHandles);
            uint tableHandle = this.responseHandles[0][getHierarchyTableResponse.OutputHandleIndex];

            PropertyTag[] properties = new PropertyTag[]
            {
                new PropertyTag()
                {
                    PropertyId = (ushort)FolderPropertyId.PidTagDisplayName,
                    PropertyType = (ushort)PropertyType.PtypString
                },
                new PropertyTag()
                {
                    PropertyId = (ushort)FolderPropertyId.PidTagFolderId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                }
            };
            List<PropertyRow> propertyRows = this.GetTableRowValue(tableHandle, (ushort)getHierarchyTableResponse.RowCount, properties);

            if (propertyRows != null)
            {
                foreach (PropertyRow propertyRow in propertyRows)
                {
                    byte[] displayNameInBytes = propertyRow.PropertyValues[0].Value;
                    string displayName = Encoding.Unicode.GetString(displayNameInBytes, 0, displayNameInBytes.Length);
                    if (displayName.Equals(folderName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        byte[] folderIdInBytes = propertyRow.PropertyValues[1].Value;
                        folderId = (ulong)BitConverter.ToInt64(folderIdInBytes, 0);
                        break;
                    }
                }
            }

            return folderId;
        }
        #endregion
    }
}