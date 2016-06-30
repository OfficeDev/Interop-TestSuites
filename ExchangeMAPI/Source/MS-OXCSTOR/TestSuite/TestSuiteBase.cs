namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// MS-OXCSTOR test suite
    /// </summary>
    [TestClass]
    [SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1401:FieldsMustBePrivate", Justification = "Disable warning SA1401 because it should not be treated like a property.")]
    public class TestSuiteBase : TestClassBase
    {
        #region Variable
        /// <summary>
        /// The instance of the server adapter
        /// </summary>
        protected IMS_OXCSTORAdapter oxcstorAdapter;

        /// <summary>
        /// The instance of the SUT control adapter
        /// </summary>
        protected IMS_OXCSTORSUTControlAdapter sutControlAdapter;

        #region Configuration variable
        /// <summary>
        /// The name of server1
        /// </summary>
        protected string server1Name;

        /// <summary>
        /// The name of server2
        /// </summary>
        protected string server2Name;

        /// <summary>
        /// The domain name
        /// </summary>
        protected string domain;

        /// <summary>
        /// The essdn of the user
        /// </summary>
        protected string userDN;

        /// <summary>
        /// The identity of public database on the server1
        /// </summary>
        protected string publicDatabaseName1;

        /// <summary>
        /// A Boolean value indicates whether second system under test exists or not.
        /// </summary>
        protected bool secondSUTExistOrNot;

        /// <summary>
        /// A Boolean value indicates whether MAPI HTTP transport is set or not.
        /// </summary>
        protected bool transportByMAPIOrNot;

        /// <summary>
        /// The user whose mailbox is on server1
        /// </summary>
        protected string userNameOfMailboxOnServer1;

        /// <summary>
        /// The password of the user  whose mailbox is on server1
        /// </summary>
        protected string passwordOfUserOfMailboxOnServer1;

        /// <summary>
        /// The essdn of the user whose mailbox is on server1
        /// </summary>
        protected string essdnOfUser1;

        /// <summary>
        /// The user whose mailbox is on server2
        /// </summary>
        protected string userNameOfMailboxOnServer2;

        /// <summary>
        /// The password of the user for MailboxOnServer2
        /// </summary>
        protected string passwordOfUserOfMailboxOnServer2;

        /// <summary>
        /// The essdn of the user whose mailbox is on server2
        /// </summary>
        protected string essdnOfUser2;
        #endregion

        /// <summary>
        /// Server object handle in the response.
        /// </summary>
        protected uint outObjHandle;

        /// <summary>
        /// Server object handle in the request
        /// </summary>
        protected uint insideObjHandle;

        /// <summary>
        /// ROP output buffer
        /// </summary>
        protected RopOutputBuffer outputBuffer;

        /// <summary>
        /// The logon request for a private mail box 
        /// </summary>
        protected RopLogonRequest logonRequestForPrivateMailBox;

        /// <summary>
        /// The logon request for public folder
        /// </summary>
        protected RopLogonRequest logonRequestForPublicFolder;

        /// <summary>
        /// Logon response
        /// </summary>
        protected RopLogonResponse logonResponse;

        /// <summary>
        /// RopGetReceiveFolder ROP request
        /// </summary>
        protected RopGetReceiveFolderRequest getReceiveFolderRequest;

        /// <summary>
        /// RopGetReceiveFolder ROP response
        /// </summary>
        protected RopGetReceiveFolderResponse getReceiveFolderResponse;

        /// <summary>
        /// RopSetReceiveFolder ROP request
        /// </summary>
        protected RopSetReceiveFolderRequest setReceiveFolderRequest;

        /// <summary>
        /// RopSetReceiveFolder ROP response
        /// </summary>
        protected RopSetReceiveFolderResponse setReceiveFolderResponse;

        /// <summary>
        /// RopGetReceiveFolderTable ROP request
        /// </summary>
        protected RopGetReceiveFolderTableRequest getReceiveFolderTableRequest;

        /// <summary>
        /// RopGetReceiveFolderTable ROP response
        /// </summary>
        protected RopGetReceiveFolderTableResponse getReceiveFolderTableResponse;

        /// <summary>
        /// RopGetStoreState ROP request
        /// </summary>
        protected RopGetStoreStateRequest getStoreStateRequest;

        /// <summary>
        /// RopGetStoreState ROP response
        /// </summary>
        protected RopGetStoreStateResponse getStoreStateResponse;

        /// <summary>
        /// RopLongTermIdFromId ROP request
        /// </summary>
        protected RopLongTermIdFromIdRequest longTermIdFromIdRequest;

        /// <summary>
        /// RopLongTermIdFromId ROP response
        /// </summary>
        protected RopLongTermIdFromIdResponse longTermIdFromIdResponse;

        /// <summary>
        /// RopGetPerUserLongTermIds ROP request
        /// </summary>
        protected RopGetPerUserLongTermIdsRequest getPerUserLongTermIdsRequest;

        /// <summary>
        /// RopGetPerUserLongTermIds ROP response
        /// </summary>
        protected RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponse;

        /// <summary>
        /// RopPublicFolderIsGhosted ROP request
        /// </summary>
        protected RopPublicFolderIsGhostedRequest publicFolderIsGhostedRequest;

        /// <summary>
        /// RopPublicFolderIsGhosted ROP response
        /// </summary>
        protected RopPublicFolderIsGhostedResponse publicFolderIsGhostedResponse;

        /// <summary>
        /// RopGetOwningServers ROP request
        /// </summary>
        protected RopGetOwningServersRequest getOwningServersRequest;

        /// <summary>
        /// RopGetOwningServers ROP response
        /// </summary>
        protected RopGetOwningServersResponse getOwningServersResponse;

        /// <summary>
        /// RopIdFromLongTermId ROP request
        /// </summary>
        protected RopIdFromLongTermIdRequest getIdFromLongTermIdRequest;

        /// <summary>
        /// RopIdFromLongTermId ROP response
        /// </summary>
        protected RopIdFromLongTermIdResponse getIdFromLongTermIdResponse;

        /// <summary>
        /// RopGetPerUserGuid ROP request
        /// </summary>
        protected RopGetPerUserGuidRequest getPerUserGuidRequest;

        /// <summary>
        /// RopGetPerUserGuid ROP Response
        /// </summary>
        protected RopGetPerUserGuidResponse getPerUserGuidResponse;

        /// <summary>
        /// RopWritePerUserInformation ROP request
        /// </summary>
        protected RopWritePerUserInformationRequest writePerUserInformationRequest;

        /// <summary>
        /// RopWritePerUserInformation ROP response
        /// </summary>
        protected RopWritePerUserInformationResponse writePerUserInformationResponse;

        /// <summary>
        /// RopReadPerUserInformation ROP request
        /// </summary>
        protected RopReadPerUserInformationRequest readPerUserInformationRequest;

        /// <summary>
        /// RopReadPerUserInformation ROP response
        /// </summary>
        protected RopReadPerUserInformationResponse readPerUserInformationResponse;

        /// <summary>
        /// Return status
        /// </summary>
        protected bool returnStatus;

        /// <summary>
        /// MessageClass "My.Class"
        /// </summary>
        protected string messageMyClass;

        /// <summary>
        /// MessageClass "My.CLASS"
        /// </summary>
        protected string messageMyClassToUpper;

        #endregion

        #region TestCase Initialization
        /// <summary>
        /// Use TestInitialize to run code before running the test in the class
        /// </summary>
        protected override void TestInitialize()
        {
            this.oxcstorAdapter = Site.GetAdapter<IMS_OXCSTORAdapter>();
            this.sutControlAdapter = Site.GetAdapter<IMS_OXCSTORSUTControlAdapter>();

            this.oxcstorAdapter.SetAutoRedirect(true);
            this.InitSUTVariable();
            this.BuildRopRequestMessage();
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            string transport = Common.GetConfigurationPropertyValue("TransportSeq", this.Site);
            if ((!this.transportByMAPIOrNot && (transport.ToLower() != "ncacn_ip_tcp")) 
                || (this.transportByMAPIOrNot && Common.IsRequirementEnabled(300000001, this.Site))
                || ((transport.ToLower() == "ncacn_ip_tcp") && Common.IsRequirementEnabled(1911, this.Site)))
            {
                switch (TestContext.TestName)
                {
                    case "MSOXCSTOR_S01_TC01_TestLogonToPrivateMailBox":
                    case "MSOXCSTOR_S01_TC10_TestOperationsOnReadWriteProperties":
                        PropertyTag propertyPidTagOutOfOfficeState;
                        TaggedPropertyValue pidTagOutOfOfficeState = new TaggedPropertyValue();

                        // According to the Open Specification MS-OXPROPS, PidTagOutOfOfficeState's id is 0x661D
                        const ushort PropertyIdForPidTagOutOfOfficeState = 0x661D;

                        propertyPidTagOutOfOfficeState.PropertyId = PropertyIdForPidTagOutOfOfficeState;

                        // According to [MS-OXPROPS], PidTagOutOfOfficeState's data type is 0x000B
                        propertyPidTagOutOfOfficeState.PropertyType = 0x000B;
                        pidTagOutOfOfficeState.PropertyTag = propertyPidTagOutOfOfficeState;
                        pidTagOutOfOfficeState.Value = new byte[1];
                        pidTagOutOfOfficeState.Value[0] = 0x00;

                        this.TrySetLogonProperty(pidTagOutOfOfficeState);
                        break;
                    case "MSOXCSTOR_S01_TC11_TestRopLogonErrorCodes":
                        if (!this.transportByMAPIOrNot)
                        {
                            this.oxcstorAdapter.DisconnectEx();

                            if (Common.IsRequirementEnabled(193, this.Site)
                                || Common.IsRequirementEnabled(1268001, this.Site))
                            {
                                string userForDisableMailbox = Common.GetConfigurationPropertyValue(ConstValues.UserForDisableMailbox, this.Site);
                                string essdn = this.sutControlAdapter.GetUserDN(this.server1Name, userForDisableMailbox);

                                // Enable the disabled mailbox
                                if (string.IsNullOrEmpty(essdn))
                                {
                                    string status = this.sutControlAdapter.EnableMailbox(userForDisableMailbox);
                                    if (status.Equals("success", StringComparison.OrdinalIgnoreCase))
                                    {
                                        int sleepSeconds = int.Parse(Common.GetConfigurationPropertyValue(ConstValues.SleepSecondsAfterEnableMailbox, this.Site));
                                        Thread.Sleep(sleepSeconds * 1000);
                                    }
                                    else
                                    {
                                        Site.Assert.Fail("The mailbox of {0} is not enabled. Error: {1}.", userForDisableMailbox, status);
                                    }
                                }
                            }
                        }

                        return;
                    case "MSOXCSTOR_S01_TC04_TestRopGetStoreState":
                        RopOpenFolderRequest openFolderRequest;
                        RopOpenFolderResponse openFolderResponse;

                        openFolderRequest.RopId = 0x02;
                        openFolderRequest.LogonId = 0x0;
                        openFolderRequest.InputHandleIndex = 0x0;
                        openFolderRequest.OutputHandleIndex = 0x01;
                        openFolderRequest.FolderId = this.logonResponse.FolderIds[4]; // Inbox folder
                        openFolderRequest.OpenModeFlags = 0x0; // opening an existing folder

                        this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

                        openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
                        Site.Assert.AreEqual<uint>(
                            0x00000000,
                            openFolderResponse.ReturnValue,
                            "0 indicates the ROP succeeds, other value indicates error occurs.");
                        uint openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];

                        RopHardDeleteMessagesAndSubfoldersRequest deleteSubfoldersOfInboxRequest = new RopHardDeleteMessagesAndSubfoldersRequest
                        {
                            RopId = 0x92,
                            InputHandleIndex = 0x00,
                            WantAsynchronous = 0x00,
                            WantDeleteAssociated = 0xff
                        };
                        this.oxcstorAdapter.DoRopCall(deleteSubfoldersOfInboxRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);

                        RopHardDeleteMessagesAndSubfoldersResponse deleteSubfoldersOfInboxResponse = (RopHardDeleteMessagesAndSubfoldersResponse)this.outputBuffer.RopsList[0];
                        Site.Assert.AreEqual<uint>(
                            0x00000000,
                            deleteSubfoldersOfInboxResponse.ReturnValue,
                            "0 indicates the ROP succeeds, other value indicates error occurs.");
                        break;
                    default:
                        break;
                }

            this.oxcstorAdapter.DisconnectEx();
            }
        }

        /// <summary>
        /// Initialize some variable
        /// </summary>
        protected void InitSUTVariable()
        {
            this.domain = Common.GetConfigurationPropertyValue(ConstValues.Domain, this.Site);

            this.server1Name = Common.GetConfigurationPropertyValue(ConstValues.Server1, this.Site);
            this.server2Name = Common.GetConfigurationPropertyValue(ConstValues.Server2, this.Site);

            this.secondSUTExistOrNot = !Common.GetConfigurationPropertyValue(ConstValues.Server2, this.Site).Equals(string.Empty);
            this.transportByMAPIOrNot = Common.GetConfigurationPropertyValue(ConstValues.TransportSeq, this.Site).ToLower().Equals("mapi_http");
            this.publicDatabaseName1 = Common.GetConfigurationPropertyValue(ConstValues.PublicDbNameOnServer1, this.Site);

            this.userNameOfMailboxOnServer1 = Common.GetConfigurationPropertyValue(ConstValues.UserNameOfMailboxOnServer1, this.Site);
            this.passwordOfUserOfMailboxOnServer1 = Common.GetConfigurationPropertyValue(ConstValues.UserPasswordOfMailboxOnServer1, this.Site);
            this.essdnOfUser1 = Common.GetConfigurationPropertyValue(ConstValues.User1ESSDN, this.Site);

            this.userNameOfMailboxOnServer2 = Common.GetConfigurationPropertyValue(ConstValues.UserNameOfMailboxOnServer2, this.Site);
            this.passwordOfUserOfMailboxOnServer2 = Common.GetConfigurationPropertyValue(ConstValues.UserPasswordOfMailboxOnServer2, this.Site);
            this.essdnOfUser2 = Common.GetConfigurationPropertyValue(ConstValues.User2ESSDN, this.Site);

            this.outObjHandle = 0;
            this.insideObjHandle = 0;
            this.returnStatus = false;
            this.userDN = Common.GetConfigurationPropertyValue(ConstValues.UserEssdn, this.Site);
            this.messageMyClass = "My.Class";
            this.messageMyClassToUpper = "MY.CLASS";
        }

        /// <summary>
        /// Build some ROP request message
        /// </summary>
        protected void BuildRopRequestMessage()
        {
            // Set value for building a logon request for a private mail box
            this.logonRequestForPrivateMailBox.RopId = (byte)RopId.RopLogon;
            this.logonRequestForPrivateMailBox.LogonId = ConstValues.LoginId;
            this.logonRequestForPrivateMailBox.OutputHandleIndex = ConstValues.OutputHandleIndex;
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private;
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;
            this.logonRequestForPrivateMailBox.StoreState = 0;
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;

            // Set value for building a logon request for a public folder
            this.logonRequestForPublicFolder.RopId = (byte)RopId.RopLogon;
            this.logonRequestForPublicFolder.LogonId = ConstValues.LoginId;
            this.logonRequestForPublicFolder.OutputHandleIndex = ConstValues.OutputHandleIndex;
            this.logonRequestForPublicFolder.LogonFlags = 0x00;
            this.logonRequestForPublicFolder.LogonFlags = (byte)LogonFlags.PublicFolder;
            this.logonRequestForPublicFolder.OpenFlags = 0x01000002;
            this.logonRequestForPublicFolder.StoreState = 0;
            this.logonRequestForPublicFolder.EssdnSize = 0;
            this.logonRequestForPublicFolder.Essdn = new byte[0];

            // Set value for building a this.getReceiveFolderRequest ROP request
            this.getReceiveFolderRequest.RopId = (byte)RopId.RopGetReceiveFolder;
            this.getReceiveFolderRequest.LogonId = ConstValues.LoginId;
            this.getReceiveFolderRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");

            // Set value for building a this.setReceiveFolderRequest ROP request
            this.setReceiveFolderRequest.RopId = (byte)RopId.RopSetReceiveFolder;
            this.setReceiveFolderRequest.LogonId = ConstValues.LoginId;
            this.setReceiveFolderRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // FolderIds[2] is Search folder, this operation is NOT supported on a search folder
            // FolderIds[0], FolderIds[1], FolderIds[3] support
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");

            // Set value for building a this.getReceiveFolderTableRequest ROP request
            this.getReceiveFolderTableRequest.RopId = (byte)RopId.RopGetReceiveFolderTable;
            this.getReceiveFolderTableRequest.LogonId = ConstValues.LoginId;
            this.getReceiveFolderTableRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a getStoreState ROP request
            this.getStoreStateRequest.RopId = (byte)RopId.RopGetStoreState;
            this.getStoreStateRequest.LogonId = ConstValues.LoginId;
            this.getStoreStateRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a getPerUserLongTermIds ROP request
            this.getPerUserLongTermIdsRequest.RopId = (byte)RopId.RopGetPerUserLongTermIds;
            this.getPerUserLongTermIdsRequest.LogonId = ConstValues.LoginId;
            this.getPerUserLongTermIdsRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a publicFolderIsGhosted ROP request
            this.publicFolderIsGhostedRequest.RopId = (byte)RopId.RopPublicFolderIsGhosted;
            this.publicFolderIsGhostedRequest.LogonId = ConstValues.LoginId;
            this.publicFolderIsGhostedRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a getOwningServers ROP request
            this.getOwningServersRequest.RopId = (byte)RopId.RopGetOwningServers;
            this.getOwningServersRequest.LogonId = ConstValues.LoginId;
            this.getOwningServersRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a longTermIdFromId Request
            this.longTermIdFromIdRequest.RopId = (byte)RopId.RopLongTermIdFromId;
            this.longTermIdFromIdRequest.LogonId = ConstValues.LoginId;
            this.longTermIdFromIdRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building an idFromLongTermId Request
            this.getIdFromLongTermIdRequest.RopId = (byte)RopId.RopIdFromLongTermId;
            this.getIdFromLongTermIdRequest.LogonId = 0x00;
            this.getIdFromLongTermIdRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building an getPerUserGuid request
            this.getPerUserGuidRequest.RopId = (byte)RopId.RopGetPerUserGuid;
            this.getPerUserGuidRequest.LogonId = ConstValues.LoginId;
            this.getPerUserGuidRequest.InputHandleIndex = ConstValues.InputHandleIndex;

            // Set value for building a writePerUserInformation request
            this.writePerUserInformationRequest.RopId = (byte)RopId.RopWritePerUserInformation;
            this.writePerUserInformationRequest.LogonId = ConstValues.LoginId;
            this.writePerUserInformationRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            this.writePerUserInformationRequest.HasFinished = 0x0;
            this.writePerUserInformationRequest.DataOffset = 0x00000000;

            // Set value for building a readPerUserInformation request
            this.readPerUserInformationRequest.RopId = (byte)RopId.RopReadPerUserInformation;
            this.readPerUserInformationRequest.LogonId = ConstValues.LoginId;
            this.readPerUserInformationRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            this.readPerUserInformationRequest.Reserved = 0x00;
            this.readPerUserInformationRequest.DataOffset = 0x00; // Set to a random number, to be changed
            this.readPerUserInformationRequest.MaxDataSize = 0x100; // maximum amount of data to be returned
        }
        #endregion

        #region Methods for operations (get/set/delete) on properties
        /// <summary>
        /// This method is used to send the ROP of RopGetPropertiesSpecificRequest to get the specific property.
        /// </summary>
        /// <param name="propTag">The specified property</param>
        /// <param name="responseFlag">When get the property successfully the responseFlag is 0, otherwise return the errorCode </param>
        /// <param name="propertyValue">When  get the property successfully the responseFlag is 0, the propertyValue is the property value, otherwise the propertyValue is 0 </param>
        /// <returns>0 indicates success, others indicates error occurs.</returns>
        protected uint GetPropertyValue(PropertyTag propTag, out uint responseFlag, out byte[] propertyValue)
        {
            uint retValue = 0;
            propertyValue = null;
            PropertyTag[] tags = new PropertyTag[1];
            responseFlag = 1;
            tags[0] = propTag;

            #region Construct the request buffer
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertySizeLimit = 0xFFFF, // Specifies the maximum size allowed for a property value returned
                PropertyTagCount = (ushort)tags.Length,
                PropertyTags = tags
            };
            #endregion

            this.oxcstorAdapter.DoRopCall(getPropertiesSpecificRequest, this.outObjHandle, ROPCommandType.RopGetPropertiesSpecific, out this.outputBuffer);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.outputBuffer.RopsList[0];
            #region Check the response

            if (getPropertiesSpecificResponse.ReturnValue == 0)
            {
                // The return value is 0 to indicate that the method invoking successfully.
                if (getPropertiesSpecificResponse.RowData.PropertyValues == null ||
                    getPropertiesSpecificResponse.RowData.PropertyValues.Count <= 0)
                {
                    // The case that the operation of set property executes failed for other reasons.
                    Site.Assert.Fail("The property value is not returned when calling RopGetPropertiesSpecific ROP.");
                }

                int lenthProperty = getPropertiesSpecificResponse.RowData.PropertyValues[0].Value.Length;
                propertyValue = new byte[lenthProperty];
                System.Array.Copy(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value, propertyValue, lenthProperty);
            }
            else
            {
                retValue = getPropertiesSpecificResponse.ReturnValue;
            }

            if (null != getPropertiesSpecificResponse.RowData)
            {
                responseFlag = getPropertiesSpecificResponse.RowData.Flag;

                if (1 == responseFlag)
                {
                    responseFlag = Common.ConvertByteArrayToUint(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
                }
            }
            #endregion

            return retValue;
        }

        /// <summary>
        /// This method is used to send the ROP of RopGetPropertiesSpecificRequest to get the specific property.
        /// </summary>
        /// <param name="propTag">The specified property</param>
        /// <param name="responseFlag">The return value</param>
        /// <returns>0 indicates success, others indicates error occurs.</returns>
        protected uint TryGetFlagLogonProperty(PropertyTag propTag, out uint responseFlag)
        {
            uint retValue = 0;
            PropertyTag[] tags = new PropertyTag[1];
            tags[0] = propTag;

            responseFlag = 1;

            #region Construct the request buffer
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertySizeLimit = 0xFFFF, // Specifies the maximum size allowed for a property value returned
                PropertyTagCount = (ushort)tags.Length,
                PropertyTags = tags
            };
            #endregion

            this.oxcstorAdapter.DoRopCall(getPropertiesSpecificRequest, this.outObjHandle, ROPCommandType.RopGetPropertiesSpecific, out this.outputBuffer);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.outputBuffer.RopsList[0];

            #region Check the response

            if (getPropertiesSpecificResponse.ReturnValue == 0)
            {
                // The return value is 0 to indicate that the method invoking successfully.
                if (getPropertiesSpecificResponse.RowData.PropertyValues == null ||
                    getPropertiesSpecificResponse.RowData.PropertyValues.Count <= 0)
                {
                    // The case that the operation of set property executes failed for other reasons.
                    Site.Assert.Fail("The property value is not returned when calling RopGetPropertiesSpecific ROP.");
                }
            }
            else
            {
                retValue = getPropertiesSpecificResponse.ReturnValue;
            }
            #endregion

            if (null != getPropertiesSpecificResponse.RowData)
            {
                responseFlag = getPropertiesSpecificResponse.RowData.Flag;

                if (1 == responseFlag)
                {
                    responseFlag = Common.ConvertByteArrayToUint(getPropertiesSpecificResponse.RowData.PropertyValues[0].Value);
                }
            }

            return retValue;
        }

        /// <summary>
        /// This method is used to send the ROP of RopGetPropertiesSpecificRequest to get the specific property.
        /// </summary>
        /// <param name="propTag">The specified property</param>
        /// <returns>0 indicates success, others indicates error occurs.</returns>
        protected uint TryGetLogonProperty(PropertyTag propTag)
        {
            uint retValue = 0;
            PropertyTag[] tags = new PropertyTag[1];
            tags[0] = propTag;

            #region Construct the request buffer
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertySizeLimit = 0xFFFF, // Specifies the maximum size allowed for a property value returned
                PropertyTagCount = (ushort)tags.Length,
                PropertyTags = tags
            };
            #endregion

            this.oxcstorAdapter.DoRopCall(getPropertiesSpecificRequest, this.outObjHandle, ROPCommandType.RopGetPropertiesSpecific, out this.outputBuffer);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.outputBuffer.RopsList[0];

            #region Check the response
            if (getPropertiesSpecificResponse.ReturnValue == 0)
            {
                // The return value is 0 to indicate that the method invoking successfully.
                if (getPropertiesSpecificResponse.RowData.PropertyValues == null ||
                    getPropertiesSpecificResponse.RowData.PropertyValues.Count <= 0)
                {
                    // The case that the operation of set property executes failed for other reasons.
                    Site.Assert.Fail("The property value is not returned when calling RopGetPropertiesSpecific ROP.");
                }
            }
            else
            {
                retValue = getPropertiesSpecificResponse.ReturnValue;
            }
            #endregion

            return retValue;
        }

        /// <summary>
        /// This method is used to send the ROP of RopSetPropertiesRequest to set the specific property.
        /// </summary>
        /// <param name="propValue">The new property value</param>
        /// <returns>0 indicates success, others indicates error occurs.</returns>
        protected uint TrySetLogonProperty(TaggedPropertyValue propValue)
        {
            uint retValue = 0;
            TaggedPropertyValue[] tags = new TaggedPropertyValue[1];
            tags[0] = propValue;

            #region Construct the request buffer
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest
            {
                RopId = 0x0A,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertyValueSize = (ushort)(tags[0].Size() + 2),
                PropertyValueCount = (ushort)tags.Length,
                PropertyValues = tags
            };

            #endregion

            this.oxcstorAdapter.DoRopCall(setPropertiesRequest, this.outObjHandle, ROPCommandType.RopSetProperties, out this.outputBuffer);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.outputBuffer.RopsList[0];

            #region Check the response
            if (setPropertiesResponse.ReturnValue == 0x00)
            {
                // The return value is 0 to indicate that the method invoking successfully.
                if (setPropertiesResponse.PropertyProblems != null
                && setPropertiesResponse.PropertyProblems.Length > 0)
                {
                    retValue = setPropertiesResponse.PropertyProblems[0].ErrorCode;
                }
            }
            else
            {
                retValue = setPropertiesResponse.ReturnValue;
            }
            #endregion

            return retValue;
        }

        /// <summary>
        /// This method is used to send the ROP of RopDeletePropertiesRequest to delete the specific property.
        /// </summary>
        /// <param name="propTag">The specified property</param>
        /// <returns>0 indicates success, others indicates error occurs.</returns>
        protected uint TryDeleteLogonProperty(PropertyTag propTag)
        {
            uint retValue = 0;
            PropertyTag[] tags = new PropertyTag[1];
            tags[0] = propTag;

            #region Construct the request buffer
            RopDeletePropertiesRequest deletePropertiesRequest;

            deletePropertiesRequest.RopId = 0x0B;
            deletePropertiesRequest.LogonId = 0x0;
            deletePropertiesRequest.InputHandleIndex = 0x0;

            // Specifies how many tags are present in PropertyTags
            deletePropertiesRequest.PropertyTagCount = (ushort)tags.Length;
            deletePropertiesRequest.PropertyTags = tags;
            #endregion

            this.oxcstorAdapter.DoRopCall(deletePropertiesRequest, this.outObjHandle, ROPCommandType.RopDeleteProperties, out this.outputBuffer);
            RopDeletePropertiesResponse deletePropertiesResponse = (RopDeletePropertiesResponse)this.outputBuffer.RopsList[0];

            #region Check the response
            if (deletePropertiesResponse.ReturnValue == 0x00)
            {
                // The return value is 0 to indicate that the method invoking successfully.
                if (deletePropertiesResponse.PropertyProblems != null && deletePropertiesResponse.PropertyProblemCount > 0)
                {
                    retValue = deletePropertiesResponse.PropertyProblems[0].ErrorCode;
                }
            }
            else
            {
                retValue = deletePropertiesResponse.ReturnValue;
            }
            #endregion

            return retValue;
        }

        /// <summary>
        /// This method is used to send the ROP of RopGetPropertiesSpecificRequest to get the specific property value.
        /// </summary>
        /// <param name="propTag">The specified property</param>
        /// <returns>Return byte array</returns>
        protected byte[] TryGetLogonPropertyValue(PropertyTag propTag)
        {
            PropertyTag[] tags = new PropertyTag[1];
            tags[0] = propTag;
            #region Construct the request buffer
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest
            {
                RopId = 0x07,
                LogonId = 0x0,
                InputHandleIndex = 0x0,
                PropertySizeLimit = 0xFFFF, // Specifies the maximum size allowed for a property value returned
                PropertyTagCount = (ushort)tags.Length,
                PropertyTags = tags
            };
            #endregion

            this.oxcstorAdapter.DoRopCall(getPropertiesSpecificRequest, this.outObjHandle, ROPCommandType.RopGetPropertiesSpecific, out this.outputBuffer);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.outputBuffer.RopsList[0];

            #region Check the response and return
            if ((getPropertiesSpecificResponse.ReturnValue == 0) && (getPropertiesSpecificResponse.RowData != null))
            {
                return getPropertiesSpecificResponse.RowData.PropertyValues[0].Value;
            }
            else
            {
                return null;
            }
            #endregion Check the response and return
        }
        #endregion

        /// <summary>
        /// Check the transport is supported
        /// </summary>
        protected void CheckTransportIsSupported()
        {
            string transport = Common.GetConfigurationPropertyValue("TransportSeq", this.Site);
            if (transport.ToLower() == "mapi_http" && !Common.IsRequirementEnabled(300000001, this.Site))
            {
                Site.Assert.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. Only Exchange 2013 SP1 supports MAPIHTTP transport.");
            }

            if (transport.ToLower() == "ncacn_ip_tcp" && !Common.IsRequirementEnabled(1911, this.Site))
            {
                Site.Assert.Inconclusive("This case can not run because ncacn_ip_tcp is not enabled to run. ncacn_ip_tcp is not supported from Exchange Server 2013.");
            }
        }

        /// <summary>
        ///  Release the folder handle
        /// </summary>
        /// <param name="objectHandle">The object handle to be released</param>
        protected void RopRelease(uint objectHandle)
        {
            RopReleaseRequest ropReleaseRequest = new RopReleaseRequest
            {
                RopId = 0x01,
                LogonId = 0x00,
                InputHandleIndex = 0x00
            };
            this.oxcstorAdapter.DoRopCall(ropReleaseRequest, objectHandle, ROPCommandType.Others, out this.outputBuffer);
        }

        /// <summary>
        /// Open a folder in the specified folder.
        /// If the folder exists, return the folder ID,
        /// If not, create the folder first, and then return the folder ID.
        /// </summary>
        /// <param name="folderName">Folder name value</param>
        /// <param name="folderType">The folder type: generic or search folder</param>
        /// <param name="parentFolderId">The parent folder ID</param>
        /// <param name="openedFolderHandle">folder handle</param>
        /// <returns>Return folder ID</returns>
        protected ulong OpenFolder(string folderName, FolderType folderType, ulong parentFolderId, out uint openedFolderHandle)
        {
            ulong folderId;

            #region Open a folder
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = parentFolderId;
            openFolderRequest.OpenModeFlags = 0x0; // Opening an existing folder

            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];

            RopHardDeleteMessagesAndSubfoldersRequest deleteSearchFolders = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                InputHandleIndex = 0,
                LogonId = 0,
                WantDeleteAssociated = 1
            };

            this.oxcstorAdapter.DoRopCall(deleteSearchFolders, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            #endregion

            #region Create a folder
            string folderNameWithNull = folderName + "\0";
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;
            createFolderRequest.RopId = 0x1C;
            createFolderRequest.LogonId = 0x0;
            createFolderRequest.InputHandleIndex = 0x0;
            createFolderRequest.OutputHandleIndex = 0x01;
            createFolderRequest.FolderType = (byte)folderType;
            createFolderRequest.UseUnicodeStrings = 0x0; // FALSE
            createFolderRequest.OpenExisting = 0xFF; // non-zero(TRUE)
            createFolderRequest.Reserved = 0x0; // FALSE
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(folderNameWithNull);
            createFolderRequest.Comment = createFolderRequest.DisplayName;
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                createFolderResponse.ReturnValue,
                "if ROP succeeds, the this.returnValue of its response is 0(success)");
            folderId = createFolderResponse.FolderId;
            #endregion

            return folderId;
        }

        /// <summary>
        /// Create the active search folder in Inbox folder.
        /// </summary>
        /// <param name="folderName">The search folder name</param>
        /// <param name="searchFolderHandle">The search folder handle</param>
        /// <param name="searchFolderId">The created search folder ID</param>
        protected void CreateSearchFolder(string folderName, out uint searchFolderHandle, out ulong searchFolderId)
        {
            #region Open Inbox folder
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = this.logonResponse.FolderIds[4]; // Inbox folder
            openFolderRequest.OpenModeFlags = 0x0; // opening an existing folder

            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            uint openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderResponse.OutputHandleIndex];

            // Create a general folder
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                OutputHandleIndex = 0x01,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0
            };

            string generalFolder = Common.GenerateResourceName(Site, "Folder") + "\0";
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(generalFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(generalFolder);
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            ulong subfolderId1 = createFolderResponse.FolderId;

            #endregion

            #region Call RopCreateFolder to create the search folder under the Inbox folder.
            string folderNameWithNull = folderName + "\0";
            createFolderRequest.FolderType = 0x02;
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(folderNameWithNull);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(folderNameWithNull);
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, createFolderResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            searchFolderHandle = this.outputBuffer.ServerObjectHandleTable[createFolderResponse.OutputHandleIndex];
            searchFolderId = createFolderResponse.FolderId;

            #endregion

            #region Call RopSetSearchCriteria to establish search criteria for created general folder.

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = 0x00,
                InputHandleIndex = 0x00
            };
            byte[] restrictData = new byte[] { 0x08, 0x1F, 0x00, 0x1A, 0x00 };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)restrictData.Length;
            setSearchCriteriaRequest.RestrictionData = restrictData;
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.RestartSearch | (uint)SetSearchFlags.ForGroundSearch | (uint)SetSearchFlags.RecursiveSearch;
            this.oxcstorAdapter.DoRopCall(setSearchCriteriaRequest, searchFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = (RopSetSearchCriteriaResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, setSearchCriteriaResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #endregion
        }

        /// <summary>
        /// Delete the specified folder
        /// </summary>
        /// <param name="folderId">The folder specified by folderId to be deleted</param>
        /// <param name="folderHandle">The folder handle</param>
        /// <returns>The returned RopDeleteFolderResponse data</returns>
        protected RopDeleteFolderResponse DeleteFolder(ulong folderId, uint folderHandle)
        {
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = 0x1D;
            deleteFolderRequest.LogonId = 0x00;
            deleteFolderRequest.InputHandleIndex = 0x00;
            deleteFolderRequest.DeleteFolderFlags = 0x01; // The folder and all of the Message objects in the folder are deleted.
            deleteFolderRequest.FolderId = folderId; // Folder to be deleted
            this.oxcstorAdapter.DoRopCall(deleteFolderRequest, folderHandle, ROPCommandType.Others, out this.outputBuffer);
            RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, deleteFolderResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            return deleteFolderResponse;
        }

        /// <summary>
        /// Create a message in public folder
        /// </summary>
        /// <param name="folderId">The ulong value</param>
        /// <param name="targetMessageHandle">The uint value</param>
        /// <returns>Return ulong value</returns>
        protected ulong CreateMessage(ulong folderId, out uint targetMessageHandle)
        {
            ulong messageId;

            #region Create a message
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            RopCreateMessageResponse createMessageResponse;
            createMessageRequest.RopId = 0x06;
            createMessageRequest.LogonId = 0x0;
            createMessageRequest.InputHandleIndex = 0x0;
            createMessageRequest.OutputHandleIndex = 0x01;
            createMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            createMessageRequest.FolderId = folderId;
            createMessageRequest.AssociatedFlag = 0x00; // NOT an FAI message
            this.oxcstorAdapter.DoRopCall(createMessageRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            createMessageResponse = (RopCreateMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                createMessageResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            targetMessageHandle = this.outputBuffer.ServerObjectHandleTable[createMessageRequest.OutputHandleIndex];
            #endregion

            #region Save the message
            RopSaveChangesMessageRequest saveChangesMessageRequest;
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageRequest.RopId = 0x0C;
            saveChangesMessageRequest.LogonId = 0x0;
            saveChangesMessageRequest.InputHandleIndex = 0x0;
            saveChangesMessageRequest.ResponseHandleIndex = 0x01;
            saveChangesMessageRequest.SaveFlags = 0x0C; // ForceSave
            this.oxcstorAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ROPCommandType.Others, out this.outputBuffer);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                saveChangesMessageResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            messageId = saveChangesMessageResponse.MessageId;
            #endregion

            return messageId;
        }

        /// <summary>
        /// Set a Message as Read
        /// </summary>
        /// <param name="folderId">The ulong value folder ID</param>
        /// <param name="messageId">The ulong value message ID</param>
        protected void SetMessageAsRead(ulong folderId, ulong messageId)
        {
            RopOpenFolderRequest openFolderRequest;

            #region Open the public folder
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = folderId; // The public folder will be opened
            openFolderRequest.OpenModeFlags = 0x0; // Opening an existing folder
            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            uint folderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];
            #endregion

            #region Set read flag
            RopSetReadFlagsRequest setReadFlagsRequest;
            setReadFlagsRequest.RopId = 0x66;
            setReadFlagsRequest.LogonId = 0x00;
            setReadFlagsRequest.InputHandleIndex = 0x0;
            setReadFlagsRequest.WantAsynchronous = 0x00; // Synchronously
            setReadFlagsRequest.ReadFlags = 0x00; // rfDefault
            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageId;
            setReadFlagsRequest.MessageIdCount = (ushort)messageIds.Length;
            setReadFlagsRequest.MessageIds = messageIds;
            this.oxcstorAdapter.DoRopCall(setReadFlagsRequest, folderHandle, ROPCommandType.Others, out this.outputBuffer);
            #endregion
        }

        /// <summary>
        /// Generate a valid IDSETWithReplGuid structure
        /// </summary>
        /// <param name="bigData">The default value is false, if the value is true, this function will generate big data which length is greater than 4096</param>
        /// <returns>Return IDSETWithReplGuid</returns>
        protected IDSETWithReplGuid GenerateRandomValidIdset(bool bigData = false)
        {
            IDSETWithReplGuid idset = new IDSETWithReplGuid
            {
                ReplGuid = this.longTermIdFromIdResponse.LongTermId.DatabaseGuid,
                Globset = new Globset()
            };
            List<GlobCnt> globCntList = new List<GlobCnt>();
            GlobCnt globcnt = new GlobCnt();

            // According to [MS-OXCFXICS] section 2.2.2.3.2. 
            // In PushCommand, Command (1 byte): A value in the range "0x01" through "0x06".
            // CommonBytes (variable): Variable length byte array to be pushed onto the common byte stack.
            // The length of the byte array is equal to the Command value ("0x01" through "0x06").
            // Here use 06 only
            PushCommand push = new PushCommand
            {
                Command = 0x06
            };
            System.Threading.Thread.Sleep(1);
            push.GenerateRandomCommandBytes();

            // The first three byte must be zero
            push.CommandBytes[0] = 0x00;
            push.CommandBytes[1] = 0x00;
            push.CommandBytes[2] = 0x00;

            globcnt.Command = push;
            globCntList.Add(globcnt);

            if (bigData)
            {
                for (int i = 0; i < 1400; i++)
                {
                    globcnt = new GlobCnt();
                    push = new PushCommand
                    {
                        Command = 0x06
                    };
                    System.Threading.Thread.Sleep(1);
                    push.GenerateRandomCommandBytes();

                    // The first three byte must be zero
                    push.CommandBytes[0] = 0x00;
                    push.CommandBytes[1] = 0x00;
                    push.CommandBytes[2] = 0x00;

                    globcnt.Command = push;
                    globCntList.Add(globcnt);
                }
            }

            GlobCnt globcnt1 = new GlobCnt();

            // In EndCommand, the value of command is "0x00".
            EndCommand end = new EndCommand();
            globcnt1.Command = end;
            globCntList.Add(globcnt1);

            idset.Globset.GlobCntList = globCntList;

            return idset;
        }

        /// <summary>
        /// Check the data in ReadPerUserInformation is valid or not
        /// </summary>
        /// <param name="data">Byte array value</param>
        /// <returns>Return Boolean value</returns>
        protected bool VerifyDataIsIDSETStructure(byte[] data)
        {
            IDSETWithReplGuid idset = new IDSETWithReplGuid();
            try
            {
                idset.Deserialize(data, 0);
            }
            catch (FormatException formatedError)
            {
                this.Site.Log.Add(LogEntryKind.Debug, "bad formatted IDSETStructure: {0}.", formatedError.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Compare two Byte Arrays are equal or not
        /// </summary>
        /// <param name="byte1">Byte array value1</param>
        /// <param name="byte2">Byte array value2</param>
        /// <returns>True if the byte arrays are equal; otherwise false</returns>
        protected bool ByteArrayEquals(byte[] byte1, byte[] byte2)
        {
            bool isEqual = false;

            if (byte1 != null && byte2 != null && (byte1.Length == byte2.Length))
            {
                int i = 0;
                for (i = 0; i < byte1.Length; i++)
                {
                    if (byte1[i] != byte2[i])
                    {
                        break;
                    }
                }

                if (i == byte1.Length)
                {
                    isEqual = true;
                }
            }
            else if ((byte1 == null) && (byte2 == null))
            {
                return true;
            }

            return isEqual;
        }

        /// <summary>
        /// Get LongTermIdFromId
        /// </summary>
        /// <param name="folderID">The folder ID</param>
        /// <returns>Return LongTermId</returns>
        protected LongTermId GetLongTermIdFromId(ulong folderID)
        {
            this.longTermIdFromIdRequest.ObjectId = folderID;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "If ROP succeeds, the this.returnValue of its response is 0 (success)");
            return this.longTermIdFromIdResponse.LongTermId;
        }

        /// <summary>
        /// Compare the 2 RopLogonResponse objects. 
        /// </summary>
        /// <param name="res1">The first RopLogonResponse object</param>
        /// <param name="res2">The second RopLogonResponse object</param>
        /// <returns>True if equal; otherwise, false</returns>
        protected bool CompareLogonResponse(RopLogonResponse res1, RopLogonResponse res2)
        {
            bool logonResponsesAreSame = this.ByteArrayEquals(res1.MailboxGuid, res2.MailboxGuid) && (res1.OutputHandleIndex == res2.OutputHandleIndex) &&
                                this.ByteArrayEquals(res1.PerUserGuid, res2.PerUserGuid) && this.ByteArrayEquals(res1.ReplGuid, res2.ReplGuid) &&
                                this.ByteArrayEquals(res1.ReplId, res2.ReplId) && (res1.ResponseFlags == res2.ResponseFlags) && (res1.ReturnValue == res2.ReturnValue) && (res1.RopId == res2.RopId) &&
                                this.ByteArrayEquals(res1.ServerName, res2.ServerName) && (res1.ServerNameSize == res2.ServerNameSize) && (res1.StoreState == res2.StoreState);
            if (!logonResponsesAreSame)
            {
                return false;
            }

            const int FolderCounts = 13;

            // If the FolderID field is null.
            if ((null != res1.FolderIds) && (null != res2.FolderIds))
            {
                // Compare every FolderId in the FolderId array.
                for (int i = 0; i < FolderCounts; i++)
                {
                    if (res1.FolderIds[i] != res2.FolderIds[i])
                    {
                        return false;
                    }
                }
            }
            else if ((null == res1.FolderIds) & (null == res2.FolderIds))
            {
                return true;
            }
            else
            {
                return false;
            }

            return true;
        }
    }
}