namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSFOLD.
    /// </summary>
    public partial class MS_OXWSFOLDAdapter : ManagedAdapterBase, IMS_OXWSFOLDAdapter
    {
        #region Verify Transport related requirements.
        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (Common.IsRequirementEnabled(3008, this.Site) && transport == TransportProtocol.HTTPS)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3008");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3008
                // When test suite running on HTTPS, if there are no exceptions or error messages returned from server, this requirement will be captured.
                Site.CaptureRequirement(
                    3008,
                    @"[In Appendix C: Product Behavior] Implementation does use secure communications via HTTPS, as defined in [RFC2818]. (Exchange Server 2007 and above follow this behavior.)");
            }
        }
        #endregion

        #region Verify soap version related requirements.

        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            this.Site.CaptureRequirement(
                3,
                @"[In Transport]The SOAP version supported is SOAP 1.1. ");
        }

        #endregion

        #region Verify all operations related requirements.
        /// <summary>
        /// Verify the requirements relate to all operations.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result."true" means success, "false" means failure.</param>
        /// <param name="responseMessage"> Common operation response.</param>
        private void VerifyAllRelatedRequirements(bool isSchemaValidated, BaseResponseMessageType responseMessage)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1036");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1036
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1036,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The type [ArrayOfResponseMessagesType] is defined as follow:
                <xs:complexType name=""ArrayOfResponseMessagesType"">
                  <xs:choice
                    maxOccurs=""unbounded""
                  >
                    <xs:element name=""CreateItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""DeleteItemResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""UpdateItemResponseMessage""
                      type=""m:UpdateItemResponseMessageType""
                     />
                    <xs:element name=""UpdateItemInRecoverableItemsResponseMessage"" 
                     type=""m:UpdateItemInRecoverableItemsResponseMessageType""/>
                    <xs:element name=""SendItemResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""DeleteFolderResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""EmptyFolderResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""CreateFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""GetFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""FindFolderResponseMessage""
                      type=""m:FindFolderResponseMessageType""
                     />
                    <xs:element name=""UpdateFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""MoveFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""CopyFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""CreateFolderPathResponseMessage"" 
                     type=""m:FolderInfoResponseMessageType""
                    />
                    <xs:element name=""CreateAttachmentResponseMessage""
                      type=""m:AttachmentInfoResponseMessageType""
                     />
                    <xs:element name=""DeleteAttachmentResponseMessage""
                      type=""m:DeleteAttachmentResponseMessageType""
                     />
                    <xs:element name=""GetAttachmentResponseMessage""
                      type=""m:AttachmentInfoResponseMessageType""
                     />
                    <xs:element name=""UploadItemsResponseMessage""
                      type=""m:UploadItemsResponseMessageType""
                     />
                    <xs:element name=""ExportItemsResponseMessage""
                      type=""m:ExportItemsResponseMessageType""
                     />
                    <xs:element name=""MarkAllItemsAsReadResponseMessage"" 
                       type=""m:ResponseMessageType""/>
                    <xs:element name=""GetClientAccessTokenResponseMessage"" 
                       type=""m:GetClientAccessTokenResponseMessageType""/>
                    <xs:element name=""GetAppManifestsResponseMessage"" type=""m:ResponseMessageType""/>
                    <xs:element name=""GetClientExtensionResponseMessage"" 
                       type=""m:ResponseMessageType""/>
                    <xs:element name=""SetClientExtensionResponseMessage"" 
                       type=""m:ResponseMessageType""/>

                    <xs:element name=""FindItemResponseMessage""
                      type=""m:FindItemResponseMessageType""
                     />
                    <xs:element name=""MoveItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""ArchiveItemResponseMessage"" type=""m:ItemInfoResponseMessageType""/>
                    <xs:element name=""CopyItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""ResolveNamesResponseMessage""
                      type=""m:ResolveNamesResponseMessageType""
                     />
                    <xs:element name=""ExpandDLResponseMessage""
                      type=""m:ExpandDLResponseMessageType""
                     />
                    <xs:element name=""GetServerTimeZonesResponseMessage""
                      type=""m:GetServerTimeZonesResponseMessageType""
                     />
                    <xs:element name=""GetEventsResponseMessage""
                      type=""m:GetEventsResponseMessageType""
                     />
                    <xs:element name=""GetStreamingEventsResponseMessage""
                      type=""m:GetStreamingEventsResponseMessageType""
                     />
                    <xs:element name=""SubscribeResponseMessage""
                      type=""m:SubscribeResponseMessageType""
                     />
                    <xs:element name=""UnsubscribeResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""SendNotificationResponseMessage""
                      type=""m:SendNotificationResponseMessageType""
                     />
                    <xs:element name=""SyncFolderHierarchyResponseMessage""
                      type=""m:SyncFolderHierarchyResponseMessageType""
                     />
                    <xs:element name=""SyncFolderItemsResponseMessage""
                      type=""m:SyncFolderItemsResponseMessageType""
                     />
                    <xs:element name=""CreateManagedFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""ConvertIdResponseMessage""
                      type=""m:ConvertIdResponseMessageType""
                     />
                    <xs:element name=""GetSharingMetadataResponseMessage""
                      type=""m:GetSharingMetadataResponseMessageType""
                     />
                    <xs:element name=""RefreshSharingFolderResponseMessage""
                      type=""m:RefreshSharingFolderResponseMessageType""
                     />
                    <xs:element name=""GetSharingFolderResponseMessage""
                      type=""m:GetSharingFolderResponseMessageType""
                     />
                    <xs:element name=""CreateUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""DeleteUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetUserConfigurationResponseMessage""
                      type=""m:GetUserConfigurationResponseMessageType""
                     />
                    <xs:element name=""UpdateUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetRoomListsResponse""
                      type=""m:GetRoomListsResponseMessageType""
                     />
                    <xs:element name=""GetRoomsResponse""
                      type=""m:GetRoomsResponseMessageType""
                     />
                      <xs:element name=""GetRemindersResponse"" 
                       type=""m:GetRemindersResponseMessageType""/>
                      <xs:element name=""PerformReminderActionResponse"" 
                       type=""m:PerformReminderActionResponseMessageType""/>
                    <xs:element name=""ApplyConversationActionResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""FindMailboxStatisticsByKeywordsResponseMessage"" type=""m:FindMailboxStatisticsByKeywordsResponseMessageType""/>
                    <xs:element name=""GetSearchableMailboxesResponseMessage"" type=""m:GetSearchableMailboxesResponseMessageType""/>
                    <xs:element name=""SearchMailboxesResponseMessage"" type=""m:SearchMailboxesResponseMessageType""/>
                    <xs:element name=""GetDiscoverySearchConfigurationResponseMessage"" type=""m:GetDiscoverySearchConfigurationResponseMessageType""/>
                    <xs:element name=""GetHoldOnMailboxesResponseMessage"" type=""m:GetHoldOnMailboxesResponseMessageType""/>
                    <xs:element name=""SetHoldOnMailboxesResponseMessage"" type=""m:SetHoldOnMailboxesResponseMessageType""/>
                      <xs:element name=""GetNonIndexableItemStatisticsResponseMessage"" type=""m:GetNonIndexableItemStatisticsResponseMessageType""/>
                      <!-- GetNonIndexableItemDetails response -->
                      <xs:element name=""GetNonIndexableItemDetailsResponseMessage"" type=""m:GetNonIndexableItemDetailsResponseMessageType""/>
                      <!-- GetUserHoldSettings response -->
                    <xs:element name=""FindPeopleResponseMessage"" type=""m:FindPeopleResponseMessageType""/>

                    <xs:element name=""GetPasswordExpirationDateResponse"" type=""m:GetPasswordExpirationDateResponseMessageType""
                    />
                      <xs:element name=""GetPersonaResponseMessage"" type=""m:GetPersonaResponseMessageType""/>
                      <xs:element name=""GetConversationItemsResponseMessage"" type=""m:GetConversationItemsResponseMessageType""/>
                      <xs:element name=""GetUserRetentionPolicyTagsResponseMessage"" type=""m:GetUserRetentionPolicyTagsResponseMessageType""/>
                      <xs:element name=""GetUserPhotoResponseMessage"" type=""m:GetUserPhotoResponseMessageType""/>
                      <xs:element name=""MarkAsJunkResponseMessage"" type=""m:MarkAsJunkResponseMessageType""/>
                  </xs:choice>
                </xs:complexType>
                ");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1094");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1094
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1094,
                @"[In m:BaseResponseMessageType Complex Type] There MUST be only one ResponseMessages element in a response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1434");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1434
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1434,
                @"[In m:ResponseMessageType Complex Type] The type [ResponseMessageType] is defined as follow:
                <xs:complexType name=""ResponseMessageType"">
                    <xs:sequence
                    minOccurs=""0""
                    >
                    <xs:element name=""MessageText""
                        type=""xs:string""
                        minOccurs=""0""
                        />
                    <xs:element name=""ResponseCode""
                        type=""m:ResponseCodeType""
                        minOccurs=""0""
                        />
                    <xs:element name=""DescriptiveLinkKey""
                        type=""xs:int""
                        minOccurs=""0""
                        />
                    <xs:element name=""MessageXml""
                        minOccurs=""0""
                    >
                        <xs:complexType>
                        <xs:sequence>
                            <xs:any
                            process_contents=""lax""
                            minOccurs=""0""
                            maxOccurs=""unbounded""
                            />
                        </xs:sequence>
                        <xs:attribute name=""ResponseClass""
                            type=""t:ResponseClassType""
                            use=""required""
                            />
                        </xs:complexType>
                    </xs:element>
                    </xs:sequence>
                </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1436");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1436
            bool isVerifiedR1436 = true;

            for (int i = 0; i < responseMessage.ResponseMessages.Items.Length; i++)
            {
                bool verifyValue = responseMessage.ResponseMessages.Items[i].ResponseClass == ResponseClassType.Error || responseMessage.ResponseMessages.Items[i].ResponseClass == ResponseClassType.Success || responseMessage.ResponseMessages.Items[i].ResponseClass == ResponseClassType.Warning;
                isVerifiedR1436 = isVerifiedR1436 && verifyValue;
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1436,
                "MS-OXWSCDATA",
                1436,
                @"[In m:ResponseMessageType Complex Type] [ResponseClass:] The following values are valid for this attribute: Success
                , Warning
                , Error.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1284");

            // Since R1436 can be captured, this requirement will be captured too.
            this.Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1284,
                @"[In m:ResponseMessageType Complex Type] This attribute [ResponseClass] MUST be present.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1339");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1339
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1339,
                @"[In t:ServerVersionInfo Element] <xs:element name=""t:ServerVersionInfo"">
                <xs:complexType>
                <xs:attribute name=""MajorVersion""
                    type=""xs:int""
                    use=""optional""
                    />
                <xs:attribute name=""MinorVersion""
                    type=""xs:int""
                    use=""optional""
                    />
                <xs:attribute name=""MajorBuildNumber""
                    type=""xs:int""
                    use=""optional""
                    />
                <xs:attribute name=""MinorBuildNumber""
                    type=""xs:int""
                    use=""optional""
                    />
                <xs:attribute name=""Version""
                    type=""xs:string""
                    use=""optional""
                    />
                </xs:complexType>
                </xs:element>");
        }

        #endregion

        #region Verify CopyFolder operation requirements.

        /// <summary>
        /// Verify requirements related to CopyFolder response.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyCopyFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R183.");

            // Verify MS-OXWSFOLD_R183.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                183,
                @"[In CopyFolder Operation]The following is the WSDL port type specification of the CopyFolder operation.
<wsdl:operation name=""CopyFolder"">
    <wsdl:input message=""tns:CopyFolderSoapIn"" />
    <wsdl:output message=""tns:CopyFolderSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R185");

            // Verify MS-OXWSFOLD_R185.
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                185,
                @"[In CopyFolder Operation][The WSDL binding of CopyFolder is defined as:]<wsdl:operation name=""CopyFolder"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CopyFolder"" />
    <wsdl:input>
        <soap:header message=""tns:CopyFolderSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:CopyFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:CopyFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""CopyFolderResult"" use=""literal"" />
        <soap:header message=""tns:CopyFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1851");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1851
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1851,
                @"[In CopyFolder Operation]The protocol client sends a CopyFolderSoapIn request WSDL message, and the protocol server responds with a CopyFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1962");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1962           
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1962,
                @"[In tns:CopyFolderSoapOut Message][The response WSDL message of CopyFolder is defined as:]<wsdl:message name=""CopyFolderSoapOut"">
    <wsdl:part name=""CopyFolderResult"" element=""tns:CopyFolderResponse"" />
    <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
</wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R197");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R197           
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                197,
                @"[In tns:CopyFolderSoapOut Message]CopyFolderResult which Element/Type is tns:CopyFolderResponse (section 3.1.4.1.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R1971");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R1971   
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1971,
                @"[In tns:CopyFolderSoapOut Message]CopyFolderResult specifies the SOAP body of the response to a CopyFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R198");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R198    
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                198,
                @"[In tns:CopyFolderSoapOut Message]ServerVersion which Element/Type is :ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R199");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R199
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                199,
                @"[In tns:CopyFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R204.");

            // Verify MS-OXWSFOLD_R204
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                204,
                @"[In CopyFolderResponse Element][The CopyFolderResponse element is defined as:]
<xs:element name=""CopyFolderResponse""
  type=""m:CopyFolderResponseType""
 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R210.");

            // Verify MS-OXWSFOLD_R210.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                210,
                @"[In m:CopyFolderResponseType Complex Type][The schema of CopyFolderResponseType is defined as:]<xs:complexType name=""CopyFolderResponseType"">
  <xs:complexContent>
    <xs:extension
      base=""m:BaseResponseMessageType""
     />
  </xs:complexContent>
</xs:complexType>");
        }
        #endregion

        #region Verify CreateFolder operation requirements.

        /// <summary>
        /// Verify requirements of CreateFolder response.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result."true" means success, "false" means failure.</param>
        /// <param name="createFolderResponse"> Response of CreateFolder.</param>
        private void VerifyCreateFolderResponse(bool isSchemaValidated, CreateFolderResponseType createFolderResponse)
        {
            FolderInfoResponseMessageType folderInfo = (FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0];

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R95.");

            // Verify MS-OXWSFOLD_R95.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                95,
                @"[In m:FolderInfoResponseMessageType Complex Type][The schema of FolderInfoResponseMessageType is defined as:]<xs:complexType name=""FolderInfoResponseMessageType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:ResponseMessageType""
                >
                    <xs:sequence>
                    <xs:element name=""Folders""
                        type=""t:ArrayOfFoldersType""
                        minOccurs=""0""
                        />
                    </xs:sequence>
                </xs:extension>
                </xs:complexContent>
            </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R220");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R220
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                220,
                @"[In CreateFolder Operation][The WSDL binding of the CreateFolder operation is defined as:]<wsdl:operation name=""CreateFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:CreateFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:CreateFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:CreateFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:header message=""tns:CreateFolderSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""CreateFolderResult"" use=""literal"" />
                    <soap:header message=""tns:CreateFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2201");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2201
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2201,
                @"[In CreateFolder Operation]The protocol client sends a CreateFolderSoapIn request WSDL message, and the protocol server responds with a CreateFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2342");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2342
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2342,
                @"[In tns:CreateFolderSoapOut Message][The response WSDL message of CreateFolder operation is defined as:]<wsdl:message name=""CreateFolderSoapOut"">
                <wsdl:part name=""CreateFolderResult"" element=""tns:CreateFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R235");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R235
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                235,
                @"[In tns:CreateFolderSoapOut Message]CreateFolderResult which Element/Type is tns:CreateFolderResponse (section 3.1.4.2.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2351");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2351
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2351,
                @"[In tns:CreateFolderSoapOut Message]CreateFolderResult specifies the SOAP body of the response to a CreateFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R236");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R236
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                236,
                @"[In tns:CreateFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R237");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R237
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                237,
                @"[In tns:CreateFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R27.");

            // Verify MS-OXWSFOLD_R27.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                27,
                @"[In t:ArrayOfFoldersType Complex Type][The schema of ArrayOfFoldersType is defined as:]
                <xs:complexType name=""ArrayOfFoldersType"">
                  <xs:choice
                    minOccurs=""0""
                    maxOccurs=""unbounded""
                  >
                    <xs:element name=""Folder""
                      type=""t:FolderType""
                     />
                    <xs:element name=""CalendarFolder""
                      type=""t:CalendarFolderType""
                     />
                    <xs:element name=""ContactsFolder""
                      type=""t:ContactsFolderType""
                     />
                    <xs:element name=""SearchFolder""
                      type=""t:SearchFolderType""
                     />
                    <xs:element name=""TasksFolder""
                      type=""t:TasksFolderType""
                     />
                  </xs:choice>
                </xs:complexType>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R218.");

            // Verify MS-OXWSFOLD_R218.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                218,
                @"[In CreateFolder Operation]The following is the WSDL port type specification of the CreateFolder operation.
                <wsdl:operation name=""CreateFolder"">
                    <wsdl:input message=""tns:CreateFolderSoapIn"" />
                    <wsdl:output message=""tns:CreateFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R242.");

            // Verify MS-OXWSFOLD_R242.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                242,
                @"[In CreateFolderResponse Element][The CreateFolderResponse element is defined as:]
                <xs:element name=""CreateFolderResponse""
                  type=""m:CreateFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R249.");

            // Verify MS-OXWSFOLD_R249.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                249,
                @"[In t:CreateFolderResponseType Complex Type][The schema of CreateFolderResponseType is defined as:]<xs:complexType name=""CreateFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R83");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R83
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                83,
                @"[In t:DistinguishedFolderIdNameType Simple Type] The DistinguishedFolderIdNameType simple type specifies well-known folders that can be referenced by name.
                <xs:simpleType name=""DistinguishedFolderIdNameType"">
                    <xs:restriction base=""xs:string"">
                        <xs:enumeration value=""calendar""/>
                        <xs:enumeration value=""contacts""/>
                        <xs:enumeration value=""deleteditems""/>
                        <xs:enumeration value=""drafts""/>
                        <xs:enumeration value=""inbox""/>
                        <xs:enumeration value=""journal""/>
                        <xs:enumeration value=""junkemail""/>
                        <xs:enumeration value=""msgfolderroot""/>
                        <xs:enumeration value=""notes""/>
                        <xs:enumeration value=""outbox""/>
                        <xs:enumeration value=""publicfoldersroot""/>
                        <xs:enumeration value=""root""/>
                        <xs:enumeration value=""searchfolders""/>
                        <xs:enumeration value=""sentitems""/>
                        <xs:enumeration value=""tasks""/>
                        <xs:enumeration value=""voicemail""/>
                        <xs:enumeration value=""recoverableitemsroot""/>
                        <xs:enumeration value=""recoverableitemsdeletions""/>
                        <xs:enumeration value=""recoverableitemsversions""/>
                        <xs:enumeration value=""recoverableitemspurges""/>
                        <xs:enumeration value=""archiveroot""/>
                        <xs:enumeration value=""archivemsgfolderroot""/>
                        <xs:enumeration value=""archivedeleteditems""/>
                        <xs:enumeration value=""archiverecoverableitemsroot""/>
                        <xs:enumeration value=""archiverecoverableitemsdeletions""/>
                        <xs:enumeration value=""archiverecoverableitemsversions""/>
                        <xs:enumeration value=""archiverecoverableitemspurges""/>
                        <xs:enumeration value=""syncissues""/>
                        <xs:enumeration value=""conflicts""/>
                        <xs:enumeration value=""localfailures""/>
                        <xs:enumeration value=""serverfailures""/>
                        <xs:enumeration value=""recipientcache""/>
                        <xs:enumeration value=""quickcontacts""/>
                        <xs:enumeration value=""conversationhistory""/>
                        <xs:enumeration value=""adminauditlogs""/>
                        <xs:enumeration value=""todosearch""/>
                        <xs:enumeration value=""mycontacts""/>
                        <xs:enumeration value=""directory"" />
                        <xs:enumeration value=""imcontactlist""/>
                        <xs:enumeration value=""peopleconnect"" />
                    </xs:restriction>
                </xs:simpleType>");

            if (folderInfo.Folders[0] is TasksFolderType)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R35");

                // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R35
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSTASK",
                    35,
                    @"[In t:TasksFolderType Complex Type] The TasksFolderType complex type specifies a Tasks folder that is contained in a mailbox.");
            }
        }
        #endregion

        #region Verify CreateManagedFolder operation requirements.

        /// <summary>
        /// Verify requirements of CreateManagedFolder operation.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyCreateManagedFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R267");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R267
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                267,
                @"[In CreateManagedFolder Operation][The WSDL binding of CreateManagedFolder operation is defined as:]<wsdl:operation name=""CreateManagedFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateManagedFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:CreateManagedFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:CreateManagedFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:CreateManagedFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""CreateManagedFolderResult"" use=""literal"" />
                    <soap:header message=""tns:CreateManagedFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2691");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2691
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2691,
                @"[In CreateManagedFolder Operation]The protocol client sends a CreateManagedFolderSoapIn request WSDL message, and the protocol server responds with a CreateManagedSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2822");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2822
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2822,
                @"[In tns:CreateManagedFolderSoapOut Message][The response WSDL message of CreateManagedFolder operation is defined as:]<wsdl:message name=""CreateManagedFolderSoapOut"">
                <wsdl:part name=""CreateManagedFolderResult"" element=""tns:CreateManagedFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R284");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R284
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                284,
                @"[In tns:CreateManagedFolderSoapOut Message]CreateManagedFolderResult which Element/Type is tns:CreateManagedFolderResponse (section 3.1.4.3.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2841");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2841
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2841,
                @"[In tns:CreateManagedFolderSoapOut Message]CreateManagedFolderResult specifies the SOAP body of the response to a CreateManagedFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R285");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R285
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                285,
                @"[In tns:CreateManagedFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R2851");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R2851
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                2851,
                @"[In tns:CreateManagedFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R265.");

            // Verify MS-OXWSFOLD_R265.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                265,
                @"[In CreateManagedFolder Operation]The following is the WSDL port type specification of the CreateManagedFolder operation.
                <wsdl:operation name=""CreateManagedFolder"">
                    <wsdl:input message=""tns:CreateManagedFolderSoapIn"" />
                    <wsdl:output message=""tns:CreateManagedFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R290.");

            // Verify MS-OXWSFOLD_R290.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                290,
                @"[In CreateManagedFolderResponse Element][The CreateManagedFolderResponse element is defined as:]
                <xs:element name=""CreateManagedFolderResponse""
                  type=""m:CreateManagedFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R301.");

            // Verify MS-OXWSFOLD_R301.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                301,
                @"[In m:CreateManagedFolderResponseType Complex Type][The schema of CreateManagedFolderResponseType is defined as:]<xs:complexType name=""CreateManagedFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");
        }
        #endregion

        #region Verify DeleteFolder operation requirements.

        /// <summary>
        /// Verify requirements of DeleteFolder operation.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyDeleteFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3102");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3102
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3102,
                @"[In DeleteFolder Operation][The WSDL binding of DeleteFolder operation is defined as:]<wsdl:operation name=""DeleteFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/DeleteFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:DeleteFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:DeleteFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:DeleteFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""DeleteFolderResult"" use=""literal"" />
                    <soap:header message=""tns:DeleteFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3103");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3103
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3103,
                @"[In DeleteFolder Operation]The protocol client sends a DeleteFolderSoapIn request WSDL message, and the protocol server MUST respond with a DeleteFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3232");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3232
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3232,
                @"[In tns:DeleteFolderSoapOut Message][The response WSDL message of DeleteFolder operation is defined as:]<wsdl:message name=""DeleteFolderSoapOut"">
                <wsdl:part name=""DeleteFolderResult"" element=""tns:DeleteFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R325");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R325
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                325,
                @"[In tns:DeleteFolderSoapOut Message]DeleteFolderResult which Element/Type is tns:DeleteFolderResponse (section 3.1.4.4.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3251");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3251
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3251,
                @"[In tns:DeleteFolderSoapOut Message]DeleteFolderResult specifies the SOAP body of the response to a DeleteFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R326");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R326
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                326,
                @"[In tns:DeleteFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R327");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R327
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                327,
                @"[In tns:DeleteFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R310.");

            // Verify MS-OXWSFOLD_R310.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                310,
                @"[In DeleteFolder Operation]The following is the WSDL port type specification of the DeleteFolder operation.
                <wsdl:operation name=""DeleteFolder"">
                    <wsdl:input message=""tns:DeleteFolderSoapIn"" />
                    <wsdl:output message=""tns:DeleteFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R332.");

            // Verify MS-OXWSFOLD_R332.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                332,
                @"[In DeleteFolderResponse Element][The DeleteFolderResponse element is defined as:]
                <xs:element name=""DeleteFolderResponse""
                  type=""m:DeleteFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R338.");

            // Verify MS-OXWSFOLD_R338.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                338,
                @"[In m:DeleteFolderResponseType Complex Type][The schema of DeleteFolderResponseType is defined as:]<xs:complexType name=""DeleteFolderResponseType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""m:BaseResponseMessageType""
                     />
                  </xs:complexContent>
                </xs:complexType>");
        }
        #endregion

        #region Verify EmptyFolder operation requirements.

        /// <summary>
        /// Verify EmptyFolder operation requirements.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyEmptyFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R347.");

            // Verify MS-OXWSFOLD_R347.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                347,
                @"[In EmptyFolder Operation]The following is the WSDL port type specification of the EmptyFolder operation.
                <wsdl:operation name=""EmptyFolder"">
                    <wsdl:input message=""tns:EmptyFolderSoapIn"" />
                    <wsdl:output message=""tns:EmptyFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3472");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3472
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3472,
                @"[In EmptyFolder Operation][The WSDL binding of EmptyFolder operation is defined as:]<wsdl:operation name=""EmptyFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/EmptyFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:EmptyFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:EmptyFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:EmptyFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""EmptyFolderResult"" use=""literal"" />
                    <soap:header message=""tns:EmptyFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3473");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3473
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3473,
                @"[In EmptyFolder Operation]The protocol client sends an EmptyFolderSoapIn request WSDL message, and the protocol server responds with an EmptyFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3612");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3612
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3612,
                @"[In tns:EmptyFolderSoapOut Message][The response WSDL message of EmptyFolder operation is defined as:]<wsdl:message name=""EmptyFolderSoapOut"">
                <wsdl:part name=""EmptyFolderResult"" element=""tns:EmptyFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R363");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R363
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                363,
                @"[In tns:EmptyFolderSoapOut Message]EmptyFolderResult which Element/Type is tns:EmptyFolderResponse (section 3.1.4.5.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3631");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3631         
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3631,
                @"[In tns:EmptyFolderSoapOut Message]EmptyFolderResult specifies the SOAP body of the response to a EmptyFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R364");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R364
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                364,
                @"[In tns:EmptyFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R365");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R365
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                365,
                @"[In tns:EmptyFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R370.");

            // Verify MS-OXWSFOLD_R370.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                370,
                @"[In EmptyFolderResponse Element][The EmptyFolderResponse element is defined as:]
                <xs:element name=""EmptyFolderResponse""
                  type=""m:EmptyFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R385.");

            // Verify MS-OXWSFOLD_R385.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                385,
                @"[In m:EmptyFolderResponseType Complex Type][The schema of EmptyFolderResponseType is defined as:]<xs:complexType name=""EmptyFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");
        }
        #endregion

        #region Verify GetFolder operation requirements.

        /// <summary>
        /// Verify GetFolder operation requirements.
        /// </summary>
        /// <param name="getFolderResponse"> Get folder response message returned from server.</param>
        /// <param name="isSchemaValidated"> A Boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyGetFolderResponse(GetFolderResponseType getFolderResponse, bool isSchemaValidated)
        {
            FolderInfoResponseMessageType allFolders = (FolderInfoResponseMessageType)getFolderResponse.ResponseMessages.Items[0];

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R386.");

            // Verify MS-OXWSFOLD_R386.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                386,
                @"[In GetFolder Operation]The following is the WSDL port type specification of the GetFolder operation.
                <wsdl:operation name=""GetFolder"">
                    <wsdl:input message=""tns:GetFolderSoapIn"" />
                    <wsdl:output message=""tns:GetFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R411.");

            // Verify MS-OXWSFOLD_R411.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                411,
                @"[In GetFolderResponse Element][The GetFolderResponse element is defined as:]
                <xs:element name=""GetFolderResponse""
                  type=""m:GetFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R418.");

            // Verify MS-OXWSFOLD_R418.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                418,
                @"[In m:GetFolderResponseType Complex Type][The schema of GetFolderResponseType is defined as:]<xs:complexType name=""GetFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");

            if (allFolders.Folders[0] is FolderType && ((FolderType)allFolders.Folders[0]).PermissionSet != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R97");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R97 
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    97,
                    @"[In t:FolderType Complex Type][The schema of FolderType is defined as:]
                    <xs:complexType name=""FolderType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:BaseFolderType""
                        >
                          <xs:sequence>
                            <xs:element name=""PermissionSet""
                              type=""t:PermissionSetType""
                              minOccurs=""0""
                             />
                            <xs:element name=""UnreadCount""
                              type=""xs:int""
                              minOccurs=""0""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");
            }

            if ((allFolders.Folders[0] is FolderType && ((FolderType)allFolders.Folders[0]).PermissionSet != null) || (allFolders.Folders[0] is ContactsFolderType && ((ContactsFolderType)allFolders.Folders[0]).PermissionSet != null) || (allFolders.Folders[0] is TasksFolderType && ((TasksFolderType)allFolders.Folders[0]).PermissionSet != null))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R143");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R143
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    143,
                    @"[In t:PermissionLevelType Simple Type][The schema of PermissionLevelType is defined as:]
                    <xs:simpleType name=""PermissionLevelType"">
                      <xs:restriction
                        base=""xs:string""
                      >
                        <xs:enumeration
                          value=""None""
                         />
                        <xs:enumeration
                          value=""Owner""
                         />
                        <xs:enumeration
                          value=""PublishingEditor""
                         />
                        <xs:enumeration
                          value=""Editor""
                         />
                        <xs:enumeration
                          value=""PublishingAuthor""
                         />
                        <xs:enumeration
                          value=""Author""
                         />
                        <xs:enumeration
                          value=""NoneditingAuthor""
                         />
                        <xs:enumeration
                          value=""Reviewer""
                         />
                        <xs:enumeration
                          value=""Contributor""
                         />
                        <xs:enumeration
                          value=""Custom""
                         />
                      </xs:restriction>
                    </xs:simpleType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R43");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R43      
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    43,
                    @"[In t:BasePermissionType Complex Type]The type of element UserId is t:UserIdType ([MS-OXWSCDATA] section 2.2.4.62).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R44");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R44
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    44,
                    @"[In t:BasePermissionType Complex Type]This element [UserId] MUST be present.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R117");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R117
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    117,
                    @"[In t:PermissionSetType Complex Type][The schema of PermissionSetType is defined as:]<xs:complexType name=""PermissionSetType"">
                <xs:sequence>
                <xs:element name=""Permissions""
                    type=""t:ArrayOfPermissionsType""
                    />
                <xs:element name=""UnknownEntries""
                    type=""t:ArrayOfUnknownEntriesType""
                    minOccurs=""0""
                    />
                </xs:sequence>
            </xs:complexType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R118");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R118
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    118,
                    @"[In t:PermissionSetType Complex Type]The type of element Permissions is t:ArrayOfPermissionsType ([MS-OXWSCDATA] section 2.2.4.9).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R121");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R121   
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    121,
                    @"[In t:PermissionType Complex Type][The schema of PermissionType is defined as:]
                <xs:complexType name=""PermissionType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""t:BasePermissionType""
                    >
                      <xs:sequence>
                        <xs:element name=""ReadItems""
                          type=""t:PermissionReadAccessType""
                          minOccurs=""0""
                          maxOccurs=""1""
                         />
                        <xs:element name=""PermissionLevel""
                          type=""t:PermissionLevelType""
                          minOccurs=""1""
                          maxOccurs=""1""
                         />
                      </xs:sequence>
                    </xs:extension>
                  </xs:complexContent>
                </xs:complexType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R123");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R123     
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    123,
                    @"[In t:PermissionType Complex Type]The type of element PermissionLevel is t:PermissionLevelType (section 2.2.5.3).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R138");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R138
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    138,
                    @"[In t:PermissionActionType Simple Type][The schema of PermissionActionType is defined as:]
                <xs:simpleType name=""PermissionActionType"">
                  <xs:restriction
                    base=""xs:string""
                  >
                    <xs:enumeration
                      value=""None""
                     />
                    <xs:enumeration
                      value=""Owned""
                     />
                    <xs:enumeration
                      value=""All""
                     />
                  </xs:restriction>
                </xs:simpleType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R155");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R155
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    155,
                    @"[In t:PermissionReadAccessType Simple Type][The schema of PermissionReadAccessType is defined as:]
                <xs:simpleType name=""PermissionReadAccessType"">
                  <xs:restriction
                    base=""xs:string""
                  >
                    <xs:enumeration
                      value=""None""
                     />
                    <xs:enumeration
                      value=""FullDetails""
                     />
                  </xs:restriction>
                </xs:simpleType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1014");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1014
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    1014,
                    @"[In t:ArrayOfPermissionsType Complex Type] The type [ArrayOfPermissionsType] is defined as follow:
                <xs:complexType name=""ArrayOfPermissionsType"">
                <xs:choice
                minOccurs=""0""
                maxOccurs=""unbounded""
                >
                <xs:element name=""Permission""
                    type=""t:PermissionType""
                    />
                </xs:choice>
            </xs:complexType>");
            }

            foreach (BaseFolderType folder in allFolders.Folders)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6603, length of folder id is:{0}", folder.FolderId.Id.Length);

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6603
                bool isVerifiedR6603 = folder.FolderId.Id.Length <= 512;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR6603,
                    6603,
                    @"[In t:BaseFolderType Complex Type]The maximum length for the FolderIdType element Id attribute is 512 bytes after base64 decoding.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6604, length of folder change key is:{0}", folder.FolderId.ChangeKey.Length);

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6604
                bool isVerifiedR6604 = folder.FolderId.ChangeKey.Length <= 512;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR6604,
                    6604,
                    @"[In t:BaseFolderType Complex Type]The maximum length for the FolerIdType ChangeKey attribute is 512 bytes after base64 decoding.");

                if (folder.ParentFolderId != null)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6803, length of parent folder id is:{0}", folder.ParentFolderId.Id.Length);

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6803
                    bool isVerifiedR6803 = folder.ParentFolderId.Id.Length <= 512;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR6803,
                        6803,
                        @"[In t:BaseFolderType Complex Type]The maximum length for the FolderIdType element Id attribute is 512 bytes after base64 decoding.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R6804, length of parent folder id is:{0}", folder.ParentFolderId.ChangeKey.Length);

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R6804
                    bool isVerifiedR6804 = folder.ParentFolderId.ChangeKey.Length <= 512;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR6804,
                        6804,
                        @"[In t:BaseFolderType Complex Type]The maximum length for the FolerIdType ChangeKey attribute is 512 bytes after base64 decoding.");
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R95");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R95
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                95,
                @"[In m:FolderInfoResponseMessageType Complex Type][The schema of FolderInfoResponseMessageType is defined as:]<xs:complexType name=""FolderInfoResponseMessageType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:ResponseMessageType""
                >
                    <xs:sequence>
                    <xs:element name=""Folders""
                        type=""t:ArrayOfFoldersType""
                        minOccurs=""0""
                        />
                    </xs:sequence>
                </xs:extension>
                </xs:complexContent>
            </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R137");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R137        
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                137,
                @"[In t:FolderClassType Simple Type][The schema of FolderClassType is defined as:]
                <xs:simpleType name=""FolderClassType"">
                  <xs:restriction
                    base=""xs:string""
                   />
                </xs:simpleType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3862");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3862
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3862,
                @"[In GetFolder Operation][The WSDL binding of GetFolder operation is defined as:]<wsdl:operation name=""GetFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:GetFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:GetFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:GetFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:header message=""tns:GetFolderSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""GetFolderResult"" use=""literal"" />
                    <soap:header message=""tns:GetFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R3863");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R3863
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                3863,
                @"[In GetFolder Operation]The protocol client sends a GetFolderSoapIn request WSDL message, and the protocol server responds with a GetFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4024");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4024
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4024,
                @"[In tns:GetFolderSoapOut Message][The response WSDL message of GetFolder operation is defined as:]<wsdl:message name=""GetFolderSoapOut"">
                <wsdl:part name=""GetFolderResult"" element=""tns:GetFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R404");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R404
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                404,
                @"[In tns:GetFolderSoapOut Message]GetFolderResult which Element/Type is tns:GetFolderResponse (section 3.1.4.6.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4041");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4041   
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4041,
                @"[In tns:GetFolderSoapOut Message]GetFolderResult specifies the SOAP body of the response to a GetManagedFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R405");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R405
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                405,
                @"[In tns:GetFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R406");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R406
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                406,
                @"[In tns:GetFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1129");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1129  
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                 "MS-OXWSCDATA",
                1129,
                @"[In t:EffectiveRightsType Complex Type] The type [EffectiveRightsType] is defined as follow:
                <xs:complexType name=""EffectiveRightsType"">
                <xs:sequence>
                <xs:element name=""CreateAssociated""
                    type=""xs:boolean""
                    />
                <xs:element name=""CreateContents""
                    type=""xs:boolean""
                    />
                <xs:element name=""CreateHierarchy""
                    type=""xs:boolean""
                    />
                <xs:element name=""Delete""
                    type=""xs:boolean""
                    />
                <xs:element name=""Modify""
                    type=""xs:boolean""
                    />
                <xs:element name=""Read""
                    type=""xs:boolean""
                    />
                <xs:element name=""ViewPrivateItems""
                    type=""xs:boolean""
                    />
                </xs:sequence>
            </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1165");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1165
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1165,
                @"[In t:FolderIdType Complex Type] The type [FolderIdType] is defined as follow:
                <xs:complexType name=""FolderIdType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""t:BaseFolderIdType""
                    >
                      <xs:attribute name=""Id""
                        type=""xs:string""
                        use=""required""
                       />
                      <xs:attribute name=""ChangeKey""
                        type=""xs:string""
                        use=""optional""
                       />
                    </xs:extension>
                  </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1297");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1297
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1297,
                @"[In t:UserIdType Complex Type] The type [UserIdType] is defined as follow:
                <xs:complexType name=""UserIdType"">
                  <xs:sequence>
                    <xs:element name=""SID""
                      type=""xs:string""
                      minOccurs=""0""
                      maxOccurs=""1""
                     />
                    <xs:element name=""PrimarySmtpAddress""
                      type=""xs:string""
                      minOccurs=""0""
                      maxOccurs=""1""
                     />
                    <xs:element name=""DisplayName""
                      type=""xs:string""
                      minOccurs=""0""
                      maxOccurs=""1""
                     />
                    <xs:element name=""DistinguishedUser""
                      type=""t:DistinguishedUserType""
                      minOccurs=""0""
                      maxOccurs=""1""
                     />
                    <xs:element name=""ExternalUserIdentity""
                      type=""xs:string""
                      minOccurs=""0""
                      maxOccurs=""1""
                     />
                  </xs:sequence>
                </xs:complexType>");

            if (allFolders.Folders[0].ManagedFolderInformation != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R102.");

                // Verify MS-OXWSFOLD_R102.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    102,
                    @"[In t:ManagedFolderInformationType Complex Type][The schema of ManagedFolderInformationType is defined as:]
                    <xs:complexType name=""ManagedFolderInformationType"">
                      <xs:sequence>
                        <xs:element name=""CanDelete""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""CanRenameOrMove""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""MustDisplayComment""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""HasQuota""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""IsManagedFoldersRoot""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""ManagedFolderId""
                          type=""xs:string""
                          minOccurs=""0""
                         />
                        <xs:element name=""Comment""
                          type=""xs:string""
                          minOccurs=""0""
                         />
                        <xs:element name=""StorageQuota""
                          type=""xs:int""
                          minOccurs=""0""
                         />
                        <xs:element name=""FolderSize""
                          type=""xs:int""
                          minOccurs=""0""
                         />
                        <xs:element name=""HomePage""
                          type=""xs:string""
                          minOccurs=""0""
                         />
                      </xs:sequence>
                    </xs:complexType>");

                if (allFolders.Folders[0].ManagedFolderInformation.CanDeleteSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R103");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R103
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        103,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element CanDelete is xs:boolean [XMLSCHEMA2].");
                }

                if (allFolders.Folders[0].ManagedFolderInformation.CanRenameOrMoveSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R105");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R105
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        105,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element CanRenameOrMove is xs:boolean.");
                }

                if (allFolders.Folders[0].ManagedFolderInformation.MustDisplayCommentSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R106");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R106
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        106,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element MustDisplayComment is xs:boolean.");
                }

                if (allFolders.Folders[0].ManagedFolderInformation.HasQuotaSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R107");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R107
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        107,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element HasQuota is xs:boolean.");
                }

                if (allFolders.Folders[0].ManagedFolderInformation.IsManagedFoldersRootSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R108");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R108
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        108,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element IsManagedFoldersRoot is xs:boolean.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R109");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R109
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    109,
                    @"[In t:ManagedFolderInformationType Complex Type]The type of element ManageFolderId is xs:string [XMLSCHEMA2].");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R111");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R111
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    111,
                    @"[In t:ManagedFolderInformationType Complex Type]The type of element Comment is xs:string.");

                if (allFolders.Folders[0].ManagedFolderInformation.StorageQuotaSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R112");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R112
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        112,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element StorageQuota is xs:int [XMLSCHEMA2].");
                }

                if (allFolders.Folders[0].ManagedFolderInformation.FolderSizeSpecified == true)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R114");

                    // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R114
                    this.Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        114,
                        @"[In t:ManagedFolderInformationType Complex Type]The type of element FolderSize is xs:int.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R79");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R79
                // All elements of ManagedFolderInformation have been verified by capture codes above (R103 to R114), so this requirement can be captured.
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    79,
                    @"[In t:BaseFolderType Complex Type]The type of element ManagedFolderInformation is t:ManagedFolderInformationType (section 2.2.4.11).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R7901");

                // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R7901
                // All elements of ManagedFolderInformation have been verified by capture codes above (R103 to R114), so this requirement can be captured.
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    7901,
                    @"[In t:BaseFolderType Complex Type]ManagedFolderInformation specifies metadata for a managed folder.");
            }
        }
        #endregion

        #region Verify MoveFolder operation requirements.

        /// <summary>
        /// Verify MoveFolder operation requirements.
        /// </summary>
        /// <param name="isSchemaValidated"> A Boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyMoveFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R431.");

            // Verify MS-OXWSFOLD_R431.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                431,
                @"[In MoveFolder Operation]The following is the WSDL port type specification of the MoveFolder operation.
                <wsdl:operation name=""MoveFolder"">
                    <wsdl:input message=""tns:MoveFolderSoapIn"" />
                    <wsdl:output message=""tns:MoveFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4312");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4312
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4312,
                @"[In MoveFolder Operation][The WSDL binding of MoveFolder operation is defined as:]<wsdl:operation name=""MoveFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MoveFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:MoveFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:MoveFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:MoveFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""MoveFolderResult"" use=""literal"" />
                    <soap:header message=""tns:MoveFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4313");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4313
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4313,
                @"[In MoveFolder Operation]The protocol client sends a MoveFolderSoapIn request WSDL message, and the protocol server responds with a MoveFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4452");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4452
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4452,
                @"[In tns:MoveFolderSoapOut Message][The response WSDL message of MoveFolder operation is defined as:]<wsdl:message name=""MoveFolderSoapOut"">
                <wsdl:part name=""MoveFolderResult"" element=""tns:MoveFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R447");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R447
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                447,
                @"[In tns:MoveFolderSoapOut Message]MoveFolderResult which Element/Type is tns:MoveFolderResponse (section 3.1.4.7.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4471");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4471
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4471,
                @"[In tns:MoveFolderSoapOut Message]MoveFolderResult specifies the SOAP body of the response to a MoveFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R448");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R448
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                448,
                @"[In tns:MoveFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R449");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R449
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                449,
                @"[In tns:MoveFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R454.");

            // Verify MS-OXWSFOLD_R454.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                454,
                @"[In MoveFolderResponse Element][The MoveFolderResponse element is defined as:]
                <xs:element name=""MoveFolderResponse""
                  type=""m:MoveFolderResponseType""
                 />");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R460.");

            // Verify MS-OXWSFOLD_R460.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                460,
                @"[In m:MoveFolderResponseType Complex Type][The schema of MoveFolderResponseType is defined as:]<xs:complexType name=""MoveFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");
        }
        #endregion

        #region Verify UpdateFolder operation requirements.

        /// <summary>
        /// Verify UpdateFolder operation requirements.
        /// </summary>
        /// <param name="isSchemaValidated"> A boolean value indicates the schema validation result. "true" means success, "false" means failure.</param>
        private void VerifyUpdateFolderResponse(bool isSchemaValidated)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R464.");

            // Verify MS-OXWSFOLD_R464.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                464,
                @"[In UpdateFolder Operation]The following is the WSDL port type specification of the UpdateFolder operation.
                <wsdl:operation name=""UpdateFolder"">
                    <wsdl:input message=""tns:UpdateFolderSoapIn"" />
                    <wsdl:output message=""tns:UpdateFolderSoapOut"" />
                </wsdl:operation>");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R489.");

            // Verify MS-OXWSFOLD_R489.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                489,
                @"[In UpdateFolderResponse Element][The UpdateFolderResponse element is defined as:]
                <xs:element name=""UpdateFolderResponse""
                  type=""m:UpdateFolderResponseType""
                 />");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4642");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4642
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4642,
                @"[In UpdateFolder Operation][The WSDL binding of UpdateFolder operation is defined as:]<wsdl:operation name=""UpdateFolder"">
                <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/UpdateFolder"" />
                <wsdl:input>
                    <soap:header message=""tns:UpdateFolderSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:UpdateFolderSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:UpdateFolderSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:header message=""tns:UpdateFolderSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                </wsdl:input>
                <wsdl:output>
                    <soap:body parts=""UpdateFolderResult"" use=""literal"" />
                    <soap:header message=""tns:UpdateFolderSoapOut"" part=""ServerVersion"" use=""literal""/>
                </wsdl:output>
            </wsdl:operation>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4643");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4643
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4643,
                @"[In UpdateFolder Operation]The protocol client sends an UpdateFolderSoapIn request WSDL message, and the protocol server responds with a UpdateFolderSoapOut response WSDL message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4812");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4812
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4812,
                @"[In tns:UpdateFolderSoapOut Message][The response WSDL message of UpdateFolder operation is defined as:]<wsdl:message name=""UpdateFolderSoapOut"">
                <wsdl:part name=""UpdateFolderResult"" element=""tns:UpdateFolderResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
            </wsdl:message>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R482");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R482
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                482,
                @"[In tns:UpdateFolderSoapOut Message]UpdateFolderResult which Element/Type is tns:UpdateFolderResponse (section 3.1.4.8.2.2).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R4821");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R4821
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                4821,
                @"[In tns:UpdateFolderSoapOut Message]UpdateFolderResult specifies the SOAP body of the response to a UpdateFolder operation request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R483");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R483  
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                483,
                @"[In tns:UpdateFolderSoapOut Message]ServerVersion which Element/Type is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R484");

            // Verify MS-OXWSFOLD requirement: MS-OXWSFOLD_R484
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                484,
                @"[In tns:UpdateFolderSoapOut Message]ServerVersion specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSFOLD_R503.");

            // Verify MS-OXWSFOLD_R503.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                503,
                @"[In m:UpdateFolderResponseType Complex Type][The schema of UpdateFolderResponseType is defined as:]<xs:complexType name=""UpdateFolderResponseType"">
                <xs:complexContent>
                <xs:extension
                    base=""m:BaseResponseMessageType""
                    />
                </xs:complexContent>
            </xs:complexType>");
        }
        #endregion
    }
}