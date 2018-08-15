namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A base class contains common methods and fields used by test cases
    /// </summary>
    [TestClass]
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Success return value constant variable.
        /// </summary>
        protected const uint Success = 0x0000;

        /// <summary>
        /// Constant string for the folder name.
        /// </summary>
        protected const string FolderName = "TemporaryFolder";

        /// <summary>
        /// Constant string for test user PidTagNormalizedSubject property.
        /// </summary>
        protected const string TestDataOfPidTagNormalizedSubject = "PidTagNormalizedSubject Test";

        /// <summary>
        /// Constant string for user name of TestUser1.
        /// </summary>
        protected const string TestUser1 = "TestUser1";

        /// <summary>
        /// Constant string for the message of ROP succeeds.
        /// </summary>
        protected const string ROPSucceedMsg = "The ROP should success and the ReturnValue of its response should be 0(success).";

        /// <summary>
        /// Null-Terminator string.
        /// </summary>
        protected const string NullTerminatorString = "\0";

        /// <summary>
        /// "@" string.
        /// </summary>
        protected const string At = "@";

        /// <summary>
        /// The common logon ID used in connection.
        /// </summary>
        protected const byte CommonLogonId = 0x00;

        /// <summary>
        /// The common InputHandleIndex used in request message.
        /// </summary>
        protected const byte CommonInputHandleIndex = 0x00;

        /// <summary>
        /// The common OutputHandleIndex used in request message.
        /// </summary>
        protected const byte CommonOutputHandleIndex = 0x01;

        /// <summary>
        /// The invalid input handle index.
        /// </summary>
        protected const byte InvalidInputHandleIndex = 0x01;

        /// <summary>
        /// Constant string for the prefix of PidTagRecipientDisplayName.
        /// </summary>
        protected const string PrefixOfDisplayName = "DisplayName-";

        /// <summary>
        /// The invalid value of a input handle. 
        /// </summary>
        protected const uint InvalidInputHandle = 0xFFFFFFFF;

        /// <summary>
        /// The time internal (millisecond) that is used by the test suite to find the message item after it sends out.
        /// </summary>
        protected int waitTime;

        /// <summary>
        /// The retry count that is used by the test suite to find the message item after it sends out.
        /// </summary>
        protected int retryCount;

        /// <summary>
        /// rawData byte
        /// </summary>
        protected byte[] rawData;

        /// <summary>
        /// Object handle
        /// </summary>
        protected uint insideObjHandle;

        /// <summary>
        /// The response object
        /// </summary>
        protected object response;

        /// <summary>
        /// The boolean value indicates that the case does not need cleanup.
        /// </summary>
        protected bool isNotNeedCleanupPrivateMailbox = false;

        /// <summary>
        /// Gets or sets an instance of IMS_OXCMSGAdapter interface.
        /// </summary>
        protected IMS_OXCMSGAdapter MSOXCMSGAdapter { get; set; }

        /// <summary>
        /// Gets or sets Message handle
        /// </summary>
        protected uint MessageHandle { get; set; }

        /// <summary>
        /// Gets or sets the response SOH list
        /// </summary>
        protected List<List<uint>> ResponseSOHs { get; set; }
        #endregion

        #region Test Case Initialization

        /// <summary>
        /// Test initialization
        /// </summary>
        protected override void TestInitialize()
        {
            // Put here code that shall be run before every test case execution
            this.MSOXCMSGAdapter = Site.GetAdapter<IMS_OXCMSGAdapter>();
            this.waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            this.retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            this.isNotNeedCleanupPrivateMailbox = false;

            // Set initial value for ref parameters response and rawData
            this.response = new object();
            this.rawData = new byte[] { };
            this.ResponseSOHs = null;
            PropertyHelper.InitializePropertyTagDic();
        }

        /// <summary>
        /// Test cleanup method
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.isNotNeedCleanupPrivateMailbox == true)
            {
                return;
            }

            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);

            // Clean inbox folder.
            this.CleanFolder(logonResponse.FolderIds[4]);

            // Clean outbox folder.
            this.CleanFolder(logonResponse.FolderIds[5]);

            // Put here code that shall be run after every test case execution
            this.rawData = null;
            this.insideObjHandle = 0; // Index of 0x00 in SOH 
            this.response = null;
            this.ResponseSOHs = null;
            this.MSOXCMSGAdapter.RpcDisconnect();
            this.MSOXCMSGAdapter.Reset();
        }

        #endregion

        #region Common Functions
        /// <summary>
        /// Check whether support MS-OXCMAPIHTTP transport.
        /// </summary>
        protected void CheckMapiHttpIsSupported()
        {
            if ((Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http") && (!Common.IsRequirementEnabled(10000, this.Site)))
            {
                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
        }

        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        protected void ConnectToServer(ConnectionType connectionType)
        {
            this.MSOXCMSGAdapter.RpcConnect(connectionType, Common.GetConfigurationPropertyValue("AdminUserName", this.Site), Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site), Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site));
        }

        /// <summary>
        /// Call RopLogon to establish a logon session between the client and the server.
        /// </summary>
        /// <param name="logonType">LogonType value</param>
        /// <param name="objHandle">Unsigned integer value</param>
        /// <returns>Return RopLogonResponse</returns>
        protected RopLogonResponse Logon(LogonType logonType, out uint objHandle)
        {
            RopLogonRequest logonRequest = new RopLogonRequest()
            {
                RopId = (byte)RopId.RopLogon, // RopId 0XFE indicates RopLogon
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                OutputHandleIndex = 0x00 // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the output Server Object is stored. 
            };
            string userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site) + "\0";

            logonRequest.StoreState = 0;

            if (LogonType.PublicFolder == logonType)
            {
                logonRequest.LogonFlags = 0x00; // Logon to public folders
                logonRequest.OpenFlags = 0x01000002;
                logonRequest.EssdnSize = 0;
                logonRequest.Essdn = null;
            }
            else
            {
                logonRequest.LogonFlags = 0x01; // Logon to a private mailbox
                logonRequest.OpenFlags = 0x01000000;
                logonRequest.EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN);
                logonRequest.Essdn = Encoding.ASCII.GetBytes(userDN);
            }
        
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(logonRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, logonResponse.ReturnValue, "Call RopLogon should success.");
            objHandle = this.ResponseSOHs[0][logonResponse.OutputHandleIndex];
            return logonResponse;
        }

        /// <summary>
        /// Get Created Message Handle
        /// </summary>
        /// <param name="folderId">The folder Id be used to create message</param>
        /// <param name="logonHandle">The RopLogon object handle</param>
        /// <returns>Return created Message Handle</returns>
        protected uint CreatedMessage(ulong folderId, uint logonHandle)
        {
            RopCreateMessageResponse createMessageResponse;
            return this.CreatedMessage(folderId, logonHandle, out createMessageResponse);
        }

        /// <summary>
        /// Get Created Message Handle
        /// </summary>
        /// <param name="folderId">The folder Id be used to create message</param>
        /// <param name="logonHandle">The RopLogon object handle</param>
        /// <param name="createMessageResponse">The response of calling RopCreateMessage</param>
        /// <returns>Return created Message Handle</returns>
        protected uint CreatedMessage(ulong folderId, uint logonHandle, out RopCreateMessageResponse createMessageResponse)
        {
            // Create a message object first
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF,
                FolderId = folderId,
                AssociatedFlag = 0x00
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createMessageRequest, logonHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createMessageResponse.ReturnValue, "Call RopCreateMessage should success.");

            // Get handle of message object, it will be used as input handle in next ROP
            uint messageHandle = this.ResponseSOHs[0][createMessageResponse.OutputHandleIndex];
            return messageHandle;
        }

        /// <summary>
        /// Release resources.
        /// </summary>
        /// <param name="handle">Unsigned integer value indicates the Server object Handle</param>
        protected void ReleaseRop(uint handle)
        {
            RopReleaseRequest relR = new RopReleaseRequest
            {
                InputHandleIndex = CommonInputHandleIndex,
                RopId = (byte)RopId.RopRelease,
                LogonId = CommonLogonId
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(relR, handle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
        }

        /// <summary>
        /// Save the attachment.
        /// </summary>
        /// <param name="objectHandle">A Server object handle.</param>
        /// <param name="saveChangesAttachmentReponse">The RopSaveChangesAttachmentResponse value.</param>
        protected void SaveAttachment(uint objectHandle, out RopSaveChangesAttachmentResponse saveChangesAttachmentReponse)
        {
            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest = new RopSaveChangesAttachmentRequest()
            {
                RopId = (byte)RopId.RopSaveChangesAttachment,
                LogonId = CommonLogonId,
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                InputHandleIndex = CommonInputHandleIndex,
                SaveFlags = (byte)SaveFlags.ForceSave
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesAttachmentRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            saveChangesAttachmentReponse = (RopSaveChangesAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesAttachmentReponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
        }

        /// <summary>
        /// Set properties' value for a message
        /// </summary>
        /// <param name="messageHandle">Unsigned integer value indicates a message object handle.</param>
        /// <param name="pts">PropertyObj list value</param>
        protected void SetPropertiesForMessage(uint messageHandle, List<PropertyObj> pts)
        {
            RopSetPropertiesResponse rpmSetResponse;
            this.SetPropertiesForMessage(messageHandle, pts, out rpmSetResponse);
        }

        /// <summary>
        /// Set properties' value for a message
        /// </summary>
        /// <param name="messageHandle">Unsigned integer value indicates a message object handle.</param>
        /// <param name="pts">PropertyObj list value</param>
        /// <param name="rpmSetResponse">The response of calling RopSetProperties.</param>
        protected void SetPropertiesForMessage(uint messageHandle, List<PropertyObj> pts, out RopSetPropertiesResponse rpmSetResponse)
        {
            List<TaggedPropertyValue> pvs = new List<TaggedPropertyValue>();

            int valueSize = 0;
            foreach (PropertyObj pt in pts)
            {
                PropertyTag pg = new PropertyTag
                {
                    PropertyId = (ushort)pt.PropertyID,
                    PropertyType = (ushort)pt.ValueTypeCode
                };

                TaggedPropertyValue pv = new TaggedPropertyValue
                {
                    PropertyTag = pg,
                    Value = (byte[])pt.Value
                };
                valueSize += pv.Size();

                pvs.Add(pv);
            }

            RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertyValueCount = (ushort)pvs.Count,
                PropertyValueSize = (ushort)(valueSize + 2),
                PropertyValues = pvs.ToArray()
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            rpmSetResponse = (RopSetPropertiesResponse)this.response;

            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, rpmSetResponse.PropertyProblemCount, "If ROP succeeds, the PropertyProblemCount of its response is 0(success).");
        }

        /// <summary>
        /// Get value of specific properties in Message object.
        /// </summary>
        /// <param name="folderId">The folder ID.</param>
        /// <param name="messageId">The message handle.</param>
        /// <param name="objectHandle">The object handle.</param>
        /// <param name="propertyList">The specific properties list.</param>
        /// <returns>A list of PropertyObj indicates property values.</returns>
        protected List<PropertyObj> GetSpecificPropertiesOfMessage(ulong folderId, ulong messageId, uint objectHandle, List<PropertyTag> propertyList)
        {
            uint messageHandle = this.OpenSpecificMessage(folderId, messageId, objectHandle, MessageOpenModeFlags.ReadOnly);

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(messageHandle, propertyList);
            this.ReleaseRop(messageHandle);

            return PropertyHelper.GetPropertyObjFromBuffer(propertyList.ToArray(), getPropertiesSpecificResponse);
        }

        /// <summary>
        /// Get value of specific properties in Message object.
        /// </summary>
        /// <param name="messageHandle">The message handle.</param>
        /// <param name="propertyList">The specific properties list.</param>
        /// <returns>The response of calling RopGetPropertiesSpecific.</returns>
        protected RopGetPropertiesSpecificResponse GetSpecificPropertiesOfMessage(uint messageHandle, List<PropertyTag> propertyList)
        {
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)propertyList.Count,
                PropertyTags = propertyList.ToArray()
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");
            return getPropertiesSpecificResponse;
        }

        /// <summary>
        /// Open a specific message.
        /// </summary>
        /// <param name="folderId">The folder Id of specific folder that contains specific message.</param>
        /// <param name="messageId">The message id of specific message.</param>
        /// <param name="objectHandle">A Server object handle.</param>
        /// <param name="openFlags">The OpenModeFlags value.</param>
        /// <returns>A Server object handle of specific message.</returns>
        protected uint OpenSpecificMessage(ulong folderId, ulong messageId, uint objectHandle, MessageOpenModeFlags openFlags)
        {
            RopOpenMessageResponse openMessageResponse;
            return this.OpenSpecificMessage(folderId, messageId, objectHandle, openFlags, out openMessageResponse);
        }

        /// <summary>
        /// Open a specific message.
        /// </summary>
        /// <param name="folderId">The folder Id of specific folder that contains specific message.</param>
        /// <param name="messageId">The message id of specific message.</param>
        /// <param name="objectHandle">A Server object handle.</param>
        /// <param name="openFlags">The OpenModeFlags value.</param>
        /// <param name="openMessageResponse">The response of RopOpenMessage.</param>
        /// <returns>A Server object handle of specific message.</returns>
        protected uint OpenSpecificMessage(ulong folderId, ulong messageId, uint objectHandle, MessageOpenModeFlags openFlags, out RopOpenMessageResponse openMessageResponse)
        {
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.

                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = folderId, // Open the message in INBOX folder in which message is created.
                OpenModeFlags = (byte)openFlags, // The message will be opened as read-only.
                MessageId = messageId // Open the saved message
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openMessageResponse = (RopOpenMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, "Call RopOpenMessage should success.");
            return this.ResponseSOHs[0][openMessageRequest.OutputHandleIndex];
        }

        /// <summary>
        /// Add recipients to specific message.
        /// </summary>
        /// <param name="recipientRows">The list of ModifyRecipientRow.</param>
        /// <param name="messageHandle">The Server object handle of Message.</param>
        /// <param name="propertyTag">The properties of Recipient</param>
        protected void AddRecipients(List<ModifyRecipientRow> recipientRows, uint messageHandle, PropertyTag[] propertyTag)
        {
            RopModifyRecipientsResponse modifyRecipientsResponse;
            this.AddRecipients(recipientRows, messageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, "Call RopModifyRecipients should success.");
        }

        /// <summary>
        /// Add recipients to specific message.
        /// </summary>
        /// <param name="recipientRows">The list of ModifyRecipientRow.</param>
        /// <param name="messageHandle">The Server object handle of Message.</param>
        /// <param name="propertyTag">The properties of Recipient</param>
        /// <param name="modifyRecipientsResponse">The response of calling RopModifyRecipients.</param>
        protected void AddRecipients(List<ModifyRecipientRow> recipientRows, uint messageHandle, PropertyTag[] propertyTag, out RopModifyRecipientsResponse modifyRecipientsResponse)
        {
            RopModifyRecipientsRequest modifyRecipientsRequest = new RopModifyRecipientsRequest
            {
                RopId = (byte)RopId.RopModifyRecipients,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                ColumnCount = Convert.ToUInt16(propertyTag.Length),
                RowCount = Convert.ToUInt16(recipientRows.Count),
                RecipientColumns = propertyTag,
                RecipientRows = recipientRows.ToArray()
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(modifyRecipientsRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)this.response;
        }

        /// <summary>
        /// Create an attachment of the given message.
        /// </summary>
        /// <param name="objectHandle">A Server object handle.</param>
        /// <param name="createAttachmentResponse">The RopCreateAttachmentResponse value.</param>
        /// <param name="attachmentId">The created attachment ID.</param>
        /// <returns>A Server object handle of the created attachment.</returns>
        protected uint CreateAttachment(uint objectHandle, out RopCreateAttachmentResponse createAttachmentResponse, out uint attachmentId)
        {
            RopCreateAttachmentRequest createAttachmentRequest = new RopCreateAttachmentRequest()
            {
                RopId = (byte)RopId.RopCreateAttachment,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createAttachmentRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            createAttachmentResponse = (RopCreateAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createAttachmentResponse.ReturnValue, "Call RopCreateAttachment should success.");
            attachmentId = createAttachmentResponse.AttachmentID;

            return this.ResponseSOHs[0][createAttachmentResponse.OutputHandleIndex];
        }

        /// <summary>
        /// Create Recipient Array 
        /// </summary>
        /// <param name="name">Recipient name value</param>
        /// <param name="rowId">RowId value</param>
        /// <returns>Return ModifyRecipientRow</returns>
        protected ModifyRecipientRow CreateModifyRecipientRow(string name, uint rowId)
        {
            return this.CreateModifyRecipientRow(name, rowId, RecipientType.PrimaryRecipient);
        }

        /// <summary>
        /// Create Recipient Array 
        /// </summary>
        /// <param name="name">Recipient name value</param>
        /// <param name="rowId">RowId value</param>
        /// <param name="recipientType">Recipient type value</param>
        /// <returns>Return ModifyRecipientRow</returns>
        protected ModifyRecipientRow CreateModifyRecipientRow(string name, uint rowId, RecipientType recipientType)
        {
            PropertyRow propertyRow = this.CreateRecipientColumns(name);

            RecipientRow recipientRow = new RecipientRow
            {
                RecipientFlags = 0x065B,
                DisplayName = Common.GetBytesFromUnicodeString(name)
            };
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            recipientRow.EmailAddress = Common.GetBytesFromUnicodeString(string.Format("{0}{1}{2}", name, TestSuiteBase.At, domainName));
            recipientRow.SimpleDisplayName = Common.GetBytesFromUnicodeString(PrefixOfDisplayName + name);
            recipientRow.RecipientColumnCount = 0x000C; // Matches ColummnCount
            recipientRow.RecipientProperties = propertyRow;

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = rowId,
                RecipientType = (byte)recipientType,
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };

            return modifyRecipientRow;
        }

        /// <summary>
        /// Create RecipientColumns
        /// </summary>
        /// <param name="name">Recipient name</param>
        /// <returns>Return PropertyRow</returns>
        protected PropertyRow CreateRecipientColumns(string name)
        {
            // Based on above recipient columns, an ARRAY of ModifyRecipientRow is created.
            // The following sample data is from MS-OXCMSG 4.7.1
            #region Configure a StandardPropertyRow: propertyRow, data is from Page 62 of MS-OXCMSG

            PropertyValue[] propertyValueArray = new PropertyValue[12];

            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006); // PidTagObjectType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000); // PidTagDisplayType
            propertyValueArray[2].Value = Common.GetBytesFromUnicodeString(name); // PidTa7BitDisplayName
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            propertyValueArray[3].Value = Common.GetBytesFromUnicodeString(string.Format("{0}{1}{2}", name, TestSuiteBase.At, domainName)); // PidTagSmtpAddress
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000); // PidTagSendInternetEncoding
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000); // PidTagDisplayTypeEx
            propertyValueArray[6].Value = Common.GetBytesFromUnicodeString(PrefixOfDisplayName + name); // PidTagRecipientDisplayName
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001); // PidTagRecipientFlags
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientTrackStatus
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientResourceState
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientOrder

            // The following sample data (0x007c and the subsequent 124(0x7c) binary)
            byte[] sampleData = 
            {
                0x7c, 0x00, 0x00, 0x00, 0x00, 0x00, 0xdc, 0xa7, 0x40,
                0xc8, 0xc0, 0x42, 0x10, 0x1a, 0xb4, 0xb9, 0x08, 0x00, 0x2b, 0x2f, 0xe1, 0x82, 0x01, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x2f, 0x6f, 0x3d, 0x46, 0x69, 0x72, 0x73, 0x74, 0x20, 0x4f, 0x72,
                0x67, 0x61, 0x6e, 0x69, 0x7a, 0x61, 0x74, 0x69, 0x6f, 0x6e, 0x2f, 0x6f, 0x75, 0x3d, 0x45, 0x78,
                0x63, 0x68, 0x61, 0x6e, 0x67, 0x65, 0x20, 0x41, 0x64, 0x6d, 0x69, 0x6e, 0x69, 0x73, 0x74, 0x72,
                0x61, 0x74, 0x69, 0x76, 0x65, 0x20, 0x47, 0x72, 0x6f, 0x75, 0x70, 0x20, 0x28, 0x46, 0x59, 0x44,
                0x49, 0x42, 0x4f, 0x48, 0x46, 0x32, 0x33, 0x53, 0x50, 0x44, 0x4c, 0x54, 0x29, 0x2f, 0x63, 0x6e,
                0x3d, 0x52, 0x65, 0x63, 0x69, 0x70, 0x69, 0x65, 0x6e, 0x74, 0x73, 0x2f, 0x63, 0x6e, 0x3d, 0x75,
                0x73, 0x65, 0x72, 0x32, 0x00
            };
            propertyValueArray[11].Value = sampleData; // PidTagRecipientEntryId

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = 0x01, PropertyValues = propertyValues
            };

            #endregion
            return propertyRow;
        }

        /// <summary>
        /// Create RecipientColumns
        /// </summary>
        /// <returns>List of PropertyTag</returns>
        protected PropertyTag[] CreateRecipientColumns()
        {
            #region recipientColumns

            // The following sample data is from MS-OXCMSG 4.7.1
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            PropertyTag tag;

            // PidTagObjectType
            tag.PropertyId = 0x0ffe;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[0] = tag;

            // PidTagDisplayType
            tag.PropertyId = 0x3900;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[1] = tag;

            // PidTagAddressBookDisplayNamePrintable
            tag.PropertyId = 0x39ff;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[2] = tag;

            // PidTagSmtpAddress
            tag.PropertyId = 0x39fe;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[3] = tag;

            // PidTagSendInternetEncoding
            tag.PropertyId = 0x3a71;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[4] = tag;

            // PidTagDisplayTypeEx
            tag.PropertyId = 0x3905;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[5] = tag;

            // PidTagRecipientDisplayName
            tag.PropertyId = 0x5ff6;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[6] = tag;

            // PidTagRecipientFlags
            tag.PropertyId = 0x5ffd;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[7] = tag;

            // PidTagRecipientTrackStatus
            tag.PropertyId = 0x5fff;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[8] = tag;

            // PidTagRecipientResourceState
            tag.PropertyId = 0x5fde;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[9] = tag;

            // PidTagRecipientOrder
            tag.PropertyId = 0x5fdf;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[10] = tag;

            // PidTagRecipientEntryId
            tag.PropertyId = 0x5ff7;
            tag.PropertyType = 0x0102; // PtypBinary
            sampleRecipientColumns[11] = tag;

            #endregion

            return sampleRecipientColumns;
        }

        /// <summary>
        /// Save Message
        /// </summary>
        /// <param name="messagesHandle">An unsigned integer value indicates a Message object handle.</param>
        /// <param name="saveFlags">A byte value indicates value of SaveFlags.</param>
        /// <returns>Return a RopSaveChangesMessageResponse object.</returns>
        protected RopSaveChangesMessageResponse SaveMessage(uint messagesHandle, byte saveFlags)
        {
            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
            {
                RopId = (byte)RopId.RopSaveChangesMessage,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                SaveFlags = saveFlags// Read write
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, messagesHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;
      
            return (RopSaveChangesMessageResponse)this.response;
        }

        /// <summary>
        /// Create an array of TaggedPropertyValue for message.
        /// </summary>
        /// <param name="size">The size of TaggedPropertyValue array.</param>
        /// <param name="valueOfPidTagAttachMethodFlags">Flags of PidTagAttachMethod property.</param>
        /// <returns>Return TaggedPropertyValue array.</returns>
        protected TaggedPropertyValue[] CreateMessageTaggedPropertyValueArrays(out int size, PidTagAttachMethodFlags valueOfPidTagAttachMethodFlags)
        {
            int arraySize = 0;
            TaggedPropertyValue[] result = new TaggedPropertyValue[1];

            result[0] = new TaggedPropertyValue
            {
                // PidTagAttachMethod
                PropertyTag = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMethod]
            };

            switch (Convert.ToInt32(valueOfPidTagAttachMethodFlags))
            {
                case 0x00000005:
                    result[0].Value = BitConverter.GetBytes(0x00000005); // 0x00000005 means Embedded Message
                    break;

                case 0x00000000:
                    result[0].Value = BitConverter.GetBytes(0x00000000); // 0x00000000 means the attachment has just been created
                    break;

                case 0x00000001:
                    result[0].Value = BitConverter.GetBytes(0x00000001); // 0x00000001 means the PidTagAttachDataBinary property contains the attachment data
                    break;

                case 0x00000002:
                    result[0].Value = BitConverter.GetBytes(0x00000002); // 0x00000002 means the PidTagAttachLongPathname containing a fully qualified path identifying the attachment to recipients with access to a common file server
                    break;

                case 0x00000004:
                    result[0].Value = BitConverter.GetBytes(0x00000004); // 0x00000004 means the PidTagAttachLongPathname containing a fully qualified path identifying the attachment
                    break;

                case 0x00000006:
                    result[0].Value = BitConverter.GetBytes(0x00000006); // 0x00000006 means the PidTagAttachDataObject contains data in an application-specific format
                    break;

                case 0x00000007:
                    result[0].Value = BitConverter.GetBytes(0x00000007); // 0x00000007 means the PidTagAttachLongPathname property contains a fully qualified path identifying the attachment. The PidNameAttachmentProviderType defines the web service API manipulating the attachment. 
                    break;

                default:
                    break;
            }

            for (int i = 0; i < result.Length; i++)
            {
                arraySize += result[i].Size();
            }

            size = arraySize;
            return result;
        }

        /// <summary>
        /// Get Opened Folder Handle
        /// </summary>
        /// <param name="folderId">The folder id be used to open folder</param>
        /// <param name="logonHandle">The RopLogon handle</param>
        /// <returns>Return created Message Handle</returns>
        protected uint OpenSpecificFolder(ulong folderId, uint logonHandle)
        {
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest()
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                FolderId = folderId,
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openFolderRequest, logonHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openFolderResponse.ReturnValue, "Call RopOpenFolder should success.");

            uint openedFolderHandle = this.ResponseSOHs[0][openFolderResponse.OutputHandleIndex];
            return openedFolderHandle;
        }

        /// <summary>
        /// Wait email be delivered to Inbox folder.
        /// </summary>
        /// <param name="title">The message subject of specified message</param>
        /// <param name="folderId">The specified folder ID.</param>
        /// <param name="logonHandle">The logon handle.</param>
        /// <param name="messageId">The value of message ID.</param>
        /// <returns>A Boolean value indicates whether the specified message has existed.</returns>
        protected bool WaitEmailBeDelivered(string title, ulong folderId, uint logonHandle, out ulong messageId)
        {
            bool isMessageReceived = false;
            int times = 0;
            messageId = 0;
            while (times < this.retryCount)
            {
                isMessageReceived = this.DetermineMessageExist(title, folderId, logonHandle, out messageId);
                if (isMessageReceived == true)
                {
                    break;
                }

                times++;
                System.Threading.Thread.Sleep(this.waitTime);
            }

            return isMessageReceived;
        }

        /// <summary>
        /// Wait email be delivered to Inbox folder.
        /// </summary>
        /// <param name="objectHandle">The Server object handle of inbox folder.</param>
        /// <param name="messageCountBeforeDelivery">The message count in inbox before this message be delivered</param>
        /// <returns>A Boolean value indicates whether the specified message has existed.</returns>
        protected bool WaitEmailBeDelivered(uint objectHandle, uint messageCountBeforeDelivery)
        {
            int times = 0;
            while (times < this.retryCount)
            {
                #region Call RopGetContentsTable to get the contents table of inbox folder before create message.
                RopGetContentsTableResponse getContentsTableResponse = this.GetContentTableSuccess(objectHandle);
                uint rowCountAfterDelivery = getContentsTableResponse.RowCount;
                uint contentTableHandle = this.ResponseSOHs[0][getContentsTableResponse.OutputHandleIndex];
                #endregion

                if (rowCountAfterDelivery == messageCountBeforeDelivery + 1)
                {
                    this.ReleaseRop(contentTableHandle);
                    return true;
                }

                this.ReleaseRop(contentTableHandle);
                System.Threading.Thread.Sleep(this.waitTime);
            }

            return false;
        }

        /// <summary>
        /// Create a subfolder under specific folder.
        /// </summary>
        /// <param name="parentFolderHandle">The Server object handle of parent folder.</param>
        /// <returns>An unsigned integer indicates the handle of the subfolder.</returns>
        protected ulong CreateSubFolder(uint parentFolderHandle)
        {
            ulong subFolderId;
            this.CreateSubFolder(parentFolderHandle, out subFolderId);
            return subFolderId;
        }

        /// <summary>
        /// Create a subfolder under specified folder.
        /// </summary>
        /// <param name="parentFolderHandle">The specified parent folder handle.</param>
        /// <param name="subFolderId">Output the subfolder Id.</param>
        /// <returns>Return the subfolder handle.</returns>
        protected uint CreateSubFolder(uint parentFolderHandle, out ulong subFolderId)
        {
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest()
            {
                RopId = (byte)RopId.RopCreateFolder, // RopId 0x1C indicates RopCreateFolder
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = 0x00,
                OpenExisting = 0x00,
                Reserved = 0x00,
                DisplayName = Encoding.ASCII.GetBytes(Common.GenerateResourceName(Site, FolderName) + TestSuiteBase.NullTerminatorString),
                Comment = Encoding.ASCII.GetBytes(FolderName + TestSuiteBase.NullTerminatorString)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createFolderRequest, parentFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createFolderResponse.ReturnValue, "Call RopCreateFolder should be success.");
            subFolderId = createFolderResponse.FolderId;

            return this.ResponseSOHs[0][createFolderResponse.OutputHandleIndex];
        }

        /// <summary>
        /// Set the specified columns that want to be returned.
        /// </summary>
        /// <param name="propertyTags">The specified columns.</param>
        /// <param name="objHandle">The object handle.</param>
        protected void SetColumnsSuccess(PropertyTag[] propertyTags, uint objHandle)
        {
            RopSetColumnsRequest setColumnsRequest = new RopSetColumnsRequest()
            {
                RopId = (byte)RopId.RopSetColumns,
                LogonId = CommonLogonId,

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle for the input Server object is stored, as specified in [MS-OXCROPS] section 2.2.5.1.1.
                InputHandleIndex = CommonInputHandleIndex,
                SetColumnsFlags = (byte)AsynchronousFlags.None,
                PropertyTagCount = (ushort)propertyTags.Length,
                PropertyTags = propertyTags
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setColumnsRequest, objHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setColumnsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
        }

        /// <summary>
        ///  Query the rows of specific table.
        /// </summary>
        /// <param name="objHandle">The handle of specific table.</param>
        /// <returns>The response of RopQueryRows.</returns>
        protected RopQueryRowsResponse QueryRowsSuccess(uint objHandle)
        {
            RopQueryRowsRequest queryRowsRequest = new RopQueryRowsRequest()
            {
                RopId = (byte)RopId.RopQueryRows,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                QueryRowsFlags = (byte)QueryRowsFlags.Advance,
                ForwardRead = 0x01,
                RowCount = 0x1000
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(queryRowsRequest, objHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, queryRowsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            return queryRowsResponse;
        }

        /// <summary>
        /// Get the content table of specific folder.
        /// </summary>
        /// <param name="objHandle">The handle of specific folder.</param>
        /// <returns>The handle of content table.</returns>
        protected RopGetContentsTableResponse GetContentTableSuccess(uint objHandle)
        {
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest()
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getContentsTableRequest, objHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetContentsTableResponse getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getContentsTableResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            return getContentsTableResponse;
        }

        /// <summary>
        /// Get the LongTermId form specified object ID.
        /// </summary>
        /// <param name="objectId">The specified object ID.</param>
        /// <param name="logonHandle">The logon handle.</param>
        /// <returns>Return the LongTermId that mapping with specified object id.</returns>
        protected LongTermId GetLongTermIdFormID(ulong objectId, uint logonHandle)
        {
            RopLongTermIdFromIdRequest longTermIdFromIdRequest = new RopLongTermIdFromIdRequest()
            {
                RopId = (byte)RopId.RopLongTermIdFromId,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set ObjectId to that got in the foregoing code, this id will be converted to a short-term ID.
                ObjectId = objectId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(longTermIdFromIdRequest, logonHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopLongTermIdFromIdResponse longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, longTermIdFromIdResponse.ReturnValue, "Call RopLongTermIdFromId should be success.");
            return longTermIdFromIdResponse.LongTermId;
        }

        /// <summary>
        /// Get the object ID form specified LongTermId.
        /// </summary>
        /// <param name="longTermID">The specified LongTermID.</param>
        /// <param name="logonHandle">The logon handle.</param>
        /// <returns>Return the object ID that mapping with specified LongTermID.</returns>
        protected ulong GetObjectIdFormLongTermID(LongTermId longTermID, uint logonHandle)
        {
            RopIdFromLongTermIdRequest ropIdFromLongTermIdRequest = new RopIdFromLongTermIdRequest()
            {
                RopId = (byte)RopId.RopIdFromLongTermId,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                LongTermId = longTermID,
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(ropIdFromLongTermIdRequest, logonHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopIdFromLongTermIdResponse ropIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, ropIdFromLongTermIdResponse.ReturnValue, "Call RopIdFromLongTermId should be success.");
            return ropIdFromLongTermIdResponse.ObjectId;
        }

        /// <summary>
        /// Set the value of properties identified by long ID or name in message.
        /// </summary>
        /// <param name="objectHandle">The specified message handle.</param>
        /// <param name="property">The PropertyName of specified property.</param>
        /// <param name="value">The value of specified property.</param>
        protected void SetNamedProperty(uint objectHandle, PropertyNameObject property, byte[] value)
        {
            #region Call RopGetPropertyIdsFromNames to get property ID.
            PropertyName[] propertyNames = new PropertyName[1];
            propertyNames[0] = property.PropertyName;

            RopGetPropertyIdsFromNamesRequest getPropertyIdsFromNamesRequest = new RopGetPropertyIdsFromNamesRequest()
            {
                RopId = (byte)RopId.RopGetPropertyIdsFromNames,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                Flags = (byte)GetPropertyIdsFromNamesFlags.Create,
                PropertyNameCount = (ushort)propertyNames.Length,
                PropertyNames = propertyNames,
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertyIdsFromNamesRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse = (RopGetPropertyIdsFromNamesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertyIdsFromNamesResponse.ReturnValue, "Call RopGetPropertyIdsFromNames should success.");
            #endregion

            #region Set property value.

            List<TaggedPropertyValue> taggedPropertyValues = new List<TaggedPropertyValue>();

            int valueSize = 0;
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = getPropertyIdsFromNamesResponse.PropertyIds[0].ID,
                PropertyType = (ushort)property.PropertyType
            };
            TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = propertyTag,
                Value = value
            };
            valueSize += taggedPropertyValue.Size();
            taggedPropertyValues.Add(taggedPropertyValue);

            RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertyValueCount = (ushort)taggedPropertyValues.Count,
                PropertyValueSize = (ushort)(valueSize + 2),
                PropertyValues = taggedPropertyValues.ToArray()
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, objectHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            RopSetPropertiesResponse rpmSetResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, rpmSetResponse.PropertyProblemCount, "If ROP succeeds, the PropertyProblemCount of its response is 0(success).");
            #endregion
        }
        #endregion

        #region Private methods for modify permission.

        /// <summary>
        /// Add a permission for a user to the permission list of the specified folder.
        /// </summary>
        /// <param name="userEssdn">UserDN used to connect server.</param>
        /// <param name="memberRights">The permission value.</param>
        /// <param name="folderHandle">The specified folder object handle.</param>
        protected void AddPermission(string userEssdn, uint memberRights, uint folderHandle)
        {
            PermissionData[] permissionsDataArray = this.GetPermissionDataArrayForAdd(userEssdn, memberRights);

            ModifyFlags modifyFlags = ModifyFlags.None;
            RopModifyPermissionsRequest modifyPermissionsRequest = this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, modifyFlags);

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(modifyPermissionsRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)this.response;
            Site.Assert.AreEqual<uint>(0, modifyPermissionsResponse.ReturnValue, "0 indicates the server adds permission successfully.");
        }

        /// <summary>
        /// Create buffer to get RopModifyPermissions request.
        /// </summary>
        /// <param name="permissionsDataArray">PermissionData array is used to set permission</param>
        /// <param name="modifyFlags">Set the ModifyFlags, specified in [MS-OXCPERM] section 2.2.2</param>
        /// <returns>A request used to modify permissions</returns>
        private RopModifyPermissionsRequest CreateModifyPermissionsRequestBuffer(PermissionData[] permissionsDataArray, ModifyFlags modifyFlags)
        {
            RopModifyPermissionsRequest modifyPermissionsRequest = new RopModifyPermissionsRequest
            {
                RopId = (byte)RopId.RopModifyPermissions,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                ModifyFlags = (byte)modifyFlags,
                ModifyCount = (ushort)permissionsDataArray.Length,
                PermissionsData = permissionsDataArray
            };

            return modifyPermissionsRequest;
        }

        /// <summary>
        /// Set the PermissionData Array for adding by the permissionUserName and permissionRight
        /// </summary>
        /// <param name="userEssdn">the ESSDN of the user</param>
        /// <param name="rights">The rights will be assigned to user</param>
        /// <returns>Return the permissionData Array for adding user</returns>
        private PermissionData[] GetPermissionDataArrayForAdd(string userEssdn, uint rights)
        {
            PropertyValue[] propertyValues = new PropertyValue[2];
            propertyValues[0] = this.CreateRightsProperty(rights);
            propertyValues[1] = this.CreateEntryIdProperty(userEssdn);

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = (byte)PermissionDataFlags.AddRow;
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;

            return permissionsDataArray;
        }

        /// <summary>
        /// Create TaggedProperty by integer rights.
        /// </summary>
        /// <param name="nrights">Integer rights</param>
        /// <returns>TaggedProperty indicate permission</returns>
        private TaggedPropertyValue CreateRightsProperty(uint nrights)
        {
            TaggedPropertyValue rightsProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = 0x6673;
            temp.PropertyType = (ushort)PropertyType.PtypInteger32;
            rightsProperty.PropertyTag = temp;
            rightsProperty.Value = BitConverter.GetBytes(nrights);

            return rightsProperty;
        }

        /// <summary>
        /// Generate the EntryId for modify the permissions.
        /// </summary>
        /// <param name="userEssdn">The user ESSDN.</param>
        /// <returns>TaggedPropertyValue indicate EntryId value.</returns>
        private TaggedPropertyValue CreateEntryIdProperty(string userEssdn)
        {
            TaggedPropertyValue entryIdProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = 0x0fff;
            temp.PropertyType = (ushort)PropertyType.PtypBinary;
            entryIdProperty.PropertyTag = temp;

            entryIdProperty.VarLength = true;
            entryIdProperty.Value = this.GetEntryId(userEssdn);
            return entryIdProperty;
        }

        /// <summary>
        /// Get EntryId by user ESSDN.
        /// </summary>
        /// <param name="userEssdn">The user ESSDN.</param>
        /// <returns>EntryId in bytes which is retrieved by server.</returns>
        private byte[] GetEntryId(string userEssdn)
        {
            // Generate the Entry ID.
            if (string.IsNullOrEmpty(userEssdn))
            {
                return new byte[0];
            }

            string distinguishedName = userEssdn + "\0";
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

        #endregion

        #region Private methods
        /// <summary>
        /// Clean the Inbox folder in private mailbox.
        /// </summary>
        /// <param name="folderID">The specified folder's FID.</param>
        private void CleanFolder(ulong folderID)
        {
            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                FolderId = folderID,
                OpenModeFlags = 0x00
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openFolderRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            uint folderHandle = this.ResponseSOHs[0][((RopOpenFolderResponse)this.response).OutputHandleIndex];

            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest()
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                WantAsynchronous = 0x00, // Synchronously
                WantDeleteAssociated = 0xFF // TRUE: delete all messages and subfolders
            };
            object temp = new object();
            this.MSOXCMSGAdapter.DoRopCall(hardDeleteMessagesAndSubfoldersRequest, folderHandle, ref temp, ref this.rawData, GetPropertiesFlags.None);
        }

        /// <summary>
        /// Determine the specified message whether has existed in specified folder.
        /// </summary>
        /// <param name="title">The message subject of specified message</param>
        /// <param name="folderId">The specified folder ID.</param>
        /// <param name="logonHandle">The logon handle.</param>
        /// <param name="messageID">The value of message ID.</param>
        /// <returns>A Boolean value indicates whether the specified message has existed.</returns>
        private bool DetermineMessageExist(string title, ulong folderId, uint logonHandle, out ulong messageID)
        {
            title = title + "\0";
            messageID = 0;

            #region Call RopOpenFolder to open INBOX folder
            uint openedFolderHandle = this.OpenSpecificFolder(folderId, logonHandle);
            #endregion

            #region Call RopGetContentsTable to get the contents table.
            RopGetContentsTableResponse ropGetContentsTableResponse = this.GetContentTableSuccess(openedFolderHandle);
            if (ropGetContentsTableResponse.RowCount == 0)
            {
                return false;
            }

            uint contentTableHandle = this.ResponseSOHs[0][ropGetContentsTableResponse.OutputHandleIndex];
            #endregion

            #region Call RopSetColumns to set the columns of table
            // Set propertyTags to a Sample ContentsTable PropertyTags created by CreateSampleContentsTablePropertyTags method.
            PropertyTag[] propertyTags = new PropertyTag[2];
            propertyTags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagNormalizedSubject];
            propertyTags[1] = new PropertyTag(0x674D, (ushort)PropertyType.PtypInteger64); // The PidTagInstID

            this.SetColumnsSuccess(propertyTags, contentTableHandle);
            #endregion

            #region Call RopQueryRows to get the rows of contents table.
            RopQueryRowsResponse queryRowsResponse = this.QueryRowsSuccess(contentTableHandle);
            #endregion

            foreach (PropertyRow row in queryRowsResponse.RowData.PropertyRows)
            {
                string actualMessageTitle = Encoding.Unicode.GetString(row.PropertyValues[0].Value);
                if (actualMessageTitle.Equals(title))
                {
                    messageID = BitConverter.ToUInt64(row.PropertyValues[1].Value, 0);
                    this.ReleaseRop(contentTableHandle);
                    this.ReleaseRop(openedFolderHandle);
                    return true;
                }
            }

            this.ReleaseRop(contentTableHandle);
            this.ReleaseRop(openedFolderHandle);
            return false;
        }
        #endregion
    }
}