namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// MS-OXCMSG protocol adapter. It is used to construct request messages and parse response message to communicate with the transport adapter. 
    /// </summary>
    public partial class MS_OXCMSGAdapter : ManagedAdapterBase, IMS_OXCMSGAdapter
    {
        /// <summary>
        /// Definition for default value of Output handle. 
        /// </summary>
        public const uint DefaultOutputHandle = 0xFFFFFFFF;

        /// <summary>
        /// The OxcropsClient instance.
        /// </summary>
        private OxcropsClient oxcropsClient;

        /// <summary>
        /// Status of connection.
        /// </summary>
        private bool isConnected;

        /// <summary>
        /// Reset the adapter.
        /// </summary>
        public override void Reset()
        {
            if (this.isConnected)
            {
                this.RpcDisconnect();
            }

            base.Reset();
        }

        #region IMS_OXCMSGAdapter members
        /// <summary>
        /// Initialize the adapter.
        /// </summary>
        /// <param name="testSite">Test site.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            PropertyHelper.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-OXCMSG";
            Common.MergeConfiguration(this.Site);
            this.oxcropsClient = new OxcropsClient(MapiContext.GetDefaultRpcContext(this.Site));
        }

        /// <summary>
        /// Connect to the server.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <param name="user">A string value indicates the domain account name that connects to server.</param>
        /// <param name="password">A string value indicates the password of the user which is used.</param>
        /// <param name="userDN">A string that identifies user who is making the EcDoConnectEx call</param>
        /// <returns>A Boolean value indicating whether connects successfully.</returns>
        public bool RpcConnect(ConnectionType connectionType, string user, string password, string userDN)
        {
            this.isConnected = this.oxcropsClient.Connect(
                    Common.GetConfigurationPropertyValue("SutComputerName", this.Site),
                    connectionType,
                    userDN,
                    Common.GetConfigurationPropertyValue("Domain", this.Site),
                    user,
                    password);

            return this.isConnected;
        }

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>Result of disconnecting.</returns>
        public bool RpcDisconnect()
        {
            bool ret = this.oxcropsClient.Disconnect();

            if (ret)
            {
                this.isConnected = false;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Send ROP request with single operation.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="getPropertiesFlag">The flag indicate the test cases expect to get which object type's properties(message's properties or attachment's properties).</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object response, ref byte[] rawData, GetPropertiesFlags getPropertiesFlag)
        {
            uint retValue;
            return this.DoRopCall(ropRequest, insideObjHandle, ref response, ref rawData, getPropertiesFlag, out retValue);
        }

        /// <summary>
        /// Send ROP request with single operation.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="getPropertiesFlag">The flag indicate the test cases expect to get which object type's properties(message's properties or attachment's properties).</param>
        /// <param name="returnValue">An unsigned integer value indicates the return value of call EcDoRpcExt2 method.</param>
        /// <returns>Server objects handles in response.</returns>
        public List<List<uint>> DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object response, ref byte[] rawData, GetPropertiesFlags getPropertiesFlag, out uint returnValue)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                ropRequest
            };

            List<uint> requestSOH = new List<uint>
            {
                insideObjHandle
            };

            if (Common.IsOutputHandleInRopRequest(ropRequest))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(DefaultOutputHandle);
            }
            
            List<IDeserializable> responseRops = new List<IDeserializable>();
            List<List<uint>> responseSOHs = new List<List<uint>>();

            // 0x10008 specifies the maximum size of the rgbOut buffer to place in Response.
            uint ret = this.oxcropsClient.RopCall(requestRops, requestSOH, ref responseRops, ref responseSOHs, ref rawData, 0x10008);
            returnValue = ret;
            if (ret == OxcRpcErrorCode.ECRpcFormat)
            {
                this.Site.Assert.Fail("Error RPC Format");
            }

            if (ret != 0)
            {
                return responseSOHs;
            }

            if (responseRops != null)
            {
                if (responseRops.Count > 0)
                {
                    response = responseRops[0];
                }
            }
            else
            {
                response = null;
            }

            if (ropRequest.GetType() == typeof(RopReleaseRequest))
            {
                return responseSOHs;
            }

            byte ropId = (byte)BitConverter.ToInt16(ropRequest.Serialize(), 0);

            List<PropertyObj> pts = null;
            switch (ropId)
            {
                case (byte)RopId.RopOpenMessage:
                    RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)response;

                    // This check is for the open specification expectation for a particular request with some valid input parameters.
                    if (openMessageResponse.ReturnValue == 0x00000000)
                    {
                        this.VerifyRopOpenMessageResponse(openMessageResponse);
                    }

                    break;

                case (byte)RopId.RopGetPropertiesSpecific:
                    // RopGetPropertiesSpecificRequest
                    pts = PropertyHelper.GetPropertyObjFromBuffer(((RopGetPropertiesSpecificRequest)ropRequest).PropertyTags, (RopGetPropertiesSpecificResponse)response);

                    foreach (PropertyObj pitem in pts)
                    {
                        // Verify capture code for MS-OXCMSG. 
                        this.VerifyMessageSyntaxDataType(pitem);
                    }

                    PropertyObj propertyObjPidTagSubjectPrefix = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagSubjectPrefix);
                    PropertyObj propertyObjPidTagNormalizedSubject = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagNormalizedSubject);

                    // Verify the message of PidTagSubjectPrefixAndPidTagNormalizedSubject
                    if (PropertyHelper.IsPropertyValid(propertyObjPidTagSubjectPrefix) || PropertyHelper.IsPropertyValid(propertyObjPidTagNormalizedSubject))
                    {
                        this.VerifyMessageSyntaxPidTagSubjectPrefixAndPidTagNormalizedSubject(propertyObjPidTagSubjectPrefix, propertyObjPidTagNormalizedSubject);
                    }

                    // Verify the requirements of PidTagAttachmentLinkId and PidTagAttachmentFlags.
                    PropertyObj pidTagAttachmentLinkId = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachmentLinkId);
                    if (PropertyHelper.IsPropertyValid(pidTagAttachmentLinkId))
                    {
                        this.VerifyMessageSyntaxPidTagAttachmentLinkIdAndPidTagAttachmentFlags(pidTagAttachmentLinkId);
                    }

                    PropertyObj pidTagAttachmentFlags = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachmentFlags);
                    if (PropertyHelper.IsPropertyValid(pidTagAttachmentFlags))
                    {
                        this.VerifyMessageSyntaxPidTagAttachmentLinkIdAndPidTagAttachmentFlags(pidTagAttachmentFlags);
                    }

                    // Verify the requirements of PidTagDisplayName
                    PropertyObj pidTagDisplayName = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagDisplayName);
                    PropertyObj pidTagAttachLongFilename = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachLongFilename);

                    if (PropertyHelper.IsPropertyValid(pidTagDisplayName) && PropertyHelper.IsPropertyValid(pidTagAttachLongFilename))
                    {
                        this.VerifyMessageSyntaxPidTagDisplayName(pidTagDisplayName, pidTagAttachLongFilename);
                    }

                    PropertyObj pidTagObjectType = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagObjectType);
                    PropertyObj pidTagRecordKey = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagRecordKey);

                    this.VerifyPidTagObjectTypeAndPidTagRecordKey(pidTagObjectType, pidTagRecordKey);
                    break;

                case (byte)RopId.RopGetPropertiesAll:
                    RopGetPropertiesAllResponse getPropertiesAllResponse = (RopGetPropertiesAllResponse)response;
                    pts = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);

                    foreach (PropertyObj pitem in pts)
                    {
                        // Verify capture code for MS-OXCMSG. 
                        this.VerifyMessageSyntaxDataType(pitem);
                    }

                    // Verify the requirements of PidTagArchiveDate
                    PropertyObj pidTagArchiveDateObj = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagArchiveDate);
                    PropertyObj pidTagStartDateEtc = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagStartDateEtc);

                    if (PropertyHelper.IsPropertyValid(pidTagArchiveDateObj))
                    {
                        if (PropertyHelper.IsPropertyValid(pidTagStartDateEtc))
                        {
                            byte[] byteDest = new byte[8];
                            Array.Copy((byte[])pidTagStartDateEtc.Value, 6, byteDest, 0, 8);
                            this.VerifyMessageSyntaxPidTagArchiveDate(pidTagArchiveDateObj, DateTime.FromFileTimeUtc(BitConverter.ToInt64(byteDest, 0)));
                        }
                    }

                    PropertyObj pidTagAccessLevel = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAccessLevel);
                    pidTagRecordKey = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagRecordKey);

                    if (getPropertiesFlag == GetPropertiesFlags.MessageProperties)
                    {
                        PropertyObj pidTagAccess = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAccess);

                        PropertyObj pidTagChangeKey = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagChangeKey);
                        PropertyObj pidTagCreationTime = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagCreationTime);
                        PropertyObj pidTagLastModificationTime = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagLastModificationTime);
                        PropertyObj pidTagLastModifierName = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagLastModifierName);
                        PropertyObj pidTagSearchKey = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagSearchKey);

                        // Verify properties PidTagAccess, PidTagAccessLevel, PidTagChangeKey, PidTagCreationTime, PidTagLastModificationTime, PidTagLastModifierName and PidTagSearchKey exist on all Message objects.
                        this.VerifyPropertiesExistOnAllMessageObject(pidTagAccess, pidTagAccessLevel, pidTagChangeKey, pidTagCreationTime, pidTagLastModificationTime, pidTagLastModifierName, pidTagSearchKey);
                    }

                    if (getPropertiesFlag == GetPropertiesFlags.AttachmentProperties)
                    {
                        // Verify properties PidTagAccessLevel and PidTagRecordKey exist on any Attachment object.
                        this.VerifyPropertiesExistOnAllAttachmentObject(pidTagAccessLevel, pidTagRecordKey);
                    }

                    break;

                case (byte)RopId.RopCreateMessage:
                    RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)response;

                    // Adapter requirements related with RopCreateMessage will be verified if the response is a successful one.
                    if (createMessageResponse.ReturnValue == 0x00000000)
                    {
                        int hasMessageId = createMessageResponse.HasMessageId;
                        this.VerifyMessageSyntaxHasMessageId(hasMessageId);
                    }

                    break;

                case (byte)RopId.RopReadRecipients:
                    RopReadRecipientsResponse readRecipientsResponse = (RopReadRecipientsResponse)response;

                    // Adapter requirements related with RopReadRecipients will be verified if the response is a successful one.
                    if (readRecipientsResponse.ReturnValue == 0x00000000)
                    {
                        this.VerifyMessageSyntaxRowCount(readRecipientsResponse);
                    }

                    break;

                case (byte)RopId.RopSetMessageStatus:
                    RopSetMessageStatusResponse setMessageStatusResponse = (RopSetMessageStatusResponse)response;

                    // Adapter requirements related with RopSetMessageStatus will be verified if the response is a successful one.
                    if (setMessageStatusResponse.ReturnValue == 0x00000000)
                    {
                        this.VerifyMessageSyntaxMessageStatusFlags(setMessageStatusResponse);
                    }

                    break;

                case (byte)RopId.RopCreateAttachment:
                    RopCreateAttachmentResponse createAttachmentResponse = (RopCreateAttachmentResponse)response;

                    // Adapter requirements related with RopCreateAttachment will be verified if the response is a successful one.
                    if (createAttachmentResponse.ReturnValue == 0x00000000)
                    {
                        int id = (int)createAttachmentResponse.AttachmentID;
                        this.VerifyDataStructureRopCreateAttachmentResponse(createAttachmentResponse, id);
                    }

                    break;

                case (byte)RopId.RopOpenEmbeddedMessage:
                    RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)response;

                    // Adapter requirements related with RopOpenEmbeddedMessage will be verified if the response is a successful one.
                    if (openEmbeddedMessageResponse.ReturnValue == 0x00000000)
                    {
                        ulong mid = openEmbeddedMessageResponse.MessageId;
                        this.VerifyDataStructureRopOpenEmbeddedMessageResponse(openEmbeddedMessageResponse, mid);
                    }

                    break;

                case (byte)RopId.RopSetMessageReadFlag:
                    RopSetMessageReadFlagResponse setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)response;

                    // Adapter requirements related with RopSetMessageReadFlag will be verified if the response is a successful one.
                    if (setMessageReadFlagResponse.ReturnValue == 0x00000000)
                    {
                        this.VerifyMessageSyntaxReadStatusChanged(setMessageReadFlagResponse, (RopSetMessageReadFlagRequest)ropRequest);
                    }

                    break;

                case (byte)RopId.RopSetReadFlags:
                    // Adapter requirements related with RopSetReadFlags will be verified if the response is a successful one.
                    if (((RopSetReadFlagsResponse)response).ReturnValue == 0x00000000)
                    {
                        this.VerifyRopSetReadFlagsResponse((RopSetReadFlagsResponse)response);
                    }

                    break;

                case (byte)RopId.RopGetMessageStatus:
                    // Adapter requirements related with RopGetMessageStatus will be verified if the response is a successful one.
                    if (((RopSetMessageStatusResponse)response).ReturnValue == 0x00000000)
                    {
                        this.VerifyGetMessageStatusResponse((RopSetMessageStatusResponse)response);
                    }

                    break;

                default:
                    break;
            }

           this.VerifyMAPITransport();

            return responseSOHs;
        }

        /// <summary>
        /// Get the named properties value of specified Message object.
        /// </summary>
        /// <param name="longIdProperties">The list of named properties</param>
        /// <param name="messageHandle">The object handle of specified Message object.</param>
        /// <returns>Returns named property values of specified Message object.</returns>
        public Dictionary<PropertyNames, byte[]> GetNamedPropertyValues(List<PropertyNameObject> longIdProperties, uint messageHandle)
        {
            object response = null;
            byte[] rawData = null;

            #region Call RopGetPropertyIdsFromNames to get property ID.
            PropertyName[] propertyNames = new PropertyName[longIdProperties.Count];

            for (int i = 0; i < longIdProperties.Count; i++)
            {
                propertyNames[i] = longIdProperties[i].PropertyName;
            }

            RopGetPropertyIdsFromNamesRequest getPropertyIdsFromNamesRequest;
            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse;
            getPropertyIdsFromNamesRequest.RopId = (byte)RopId.RopGetPropertyIdsFromNames;
            getPropertyIdsFromNamesRequest.LogonId = 0x00;
            getPropertyIdsFromNamesRequest.InputHandleIndex = 0x00;
            getPropertyIdsFromNamesRequest.Flags = (byte)GetPropertyIdsFromNamesFlags.Create;
            getPropertyIdsFromNamesRequest.PropertyNameCount = (ushort)propertyNames.Length;
            getPropertyIdsFromNamesRequest.PropertyNames = propertyNames;

            this.DoRopCall(getPropertyIdsFromNamesRequest, messageHandle, ref response, ref rawData, GetPropertiesFlags.None);
            getPropertyIdsFromNamesResponse = (RopGetPropertyIdsFromNamesResponse)response;
            Site.Assert.AreEqual<uint>(0, getPropertyIdsFromNamesResponse.ReturnValue, "Call RopGetPropertyIdsFromNames should success.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of specific message object.
            // Get specific property for created message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest();
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = 0x00;
            getPropertiesSpecificRequest.InputHandleIndex = 0x00;
            getPropertiesSpecificRequest.PropertySizeLimit = 0xFFFF;

            PropertyTag[] tagArray = new PropertyTag[longIdProperties.Count];
            for (int j = 0; j < getPropertyIdsFromNamesResponse.PropertyIds.Length; j++)
            {
                tagArray[j] = new PropertyTag
                {
                    PropertyId = getPropertyIdsFromNamesResponse.PropertyIds[j].ID,
                    PropertyType = (ushort)longIdProperties[j].PropertyType
                };
            }

            getPropertiesSpecificRequest.PropertyTagCount = (ushort)tagArray.Length;
            getPropertiesSpecificRequest.PropertyTags = tagArray;

            this.DoRopCall(getPropertiesSpecificRequest, messageHandle, ref response, ref rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)response;
            Site.Assert.AreEqual<uint>(0, getPropertiesSpecificResponse.ReturnValue, "Calling RopGetPropertiesSpecific should be successful.");

            Dictionary<PropertyNames, byte[]> propertyList = new Dictionary<PropertyNames, byte[]>();
            PropertyObj propertyObjPidLidCommonStart = null;
            PropertyObj propertyObjPidLidCommonEnd = null;

            for (int i = 0; i < getPropertiesSpecificResponse.RowData.PropertyValues.Count; i++)
            {
                PropertyObj propertyObj = new PropertyObj
                {
                    PropertyName = longIdProperties[i].DisplayName,
                    ValueType = longIdProperties[i].PropertyType
                };
                PropertyHelper.GetPropertyObjFromBuffer(propertyObj, getPropertiesSpecificResponse.RowData.PropertyValues[i].Value);

                // Verify requirements related with named properties PidNameKeywords, PidNameContentBase, PidNameAcceptLanguage and PidNameContentClass.
                this.VerifyMessageSyntaxDataType(propertyObj);

                if (propertyObj.PropertyName == PropertyNames.PidLidCommonStart)
                {
                    propertyObjPidLidCommonStart = propertyObj;
                }

                if (propertyObj.PropertyName == PropertyNames.PidLidCommonEnd)
                {
                    propertyObjPidLidCommonEnd = propertyObj;
                }

                propertyList.Add(longIdProperties[i].DisplayName, getPropertiesSpecificResponse.RowData.PropertyValues[i].Value);
            }

            // Verify the requirements of PidLidCommonStart and PidLidCommonEnd.
            if (PropertyHelper.IsPropertyValid(propertyObjPidLidCommonStart) || PropertyHelper.IsPropertyValid(propertyObjPidLidCommonEnd))
            {
                this.VerifyMessageSyntaxPidLidCommonStartAndPidLidCommonEnd(propertyObjPidLidCommonStart, propertyObjPidLidCommonEnd);
            }
            #endregion

            return propertyList;
        }
        #endregion
    }
}