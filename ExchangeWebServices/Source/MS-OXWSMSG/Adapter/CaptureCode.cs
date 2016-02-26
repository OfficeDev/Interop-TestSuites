namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSMSG.
    /// </summary>
    public partial class MS_OXWSMSGAdapter
    {
        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R1");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R1
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                1,
                @"[In Transport] The SOAP version supported is SOAP 1.1.");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (transport == TransportProtocol.HTTPS && Common.IsRequirementEnabled(183, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R183");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R183
                // Because Adapter uses SOAP and HTTPS to communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    183,
                    @"[In Appendix C: Product Behavior] Implementation does uses secure communications via HTTPS, as defined in [RFC2818]. (The Exchange 2007 and above follow this behavior.)");
            }

            if (transport == TransportProtocol.HTTP)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R3001");

                 // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R3001
                // Because Adapter uses HTTP communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    3001,
                    @"[In Transport] The protocol MUST support SOAP over HTTP, as specified in [RFC2616].");
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the CopyItem operation and CopyItemResponseType structure. 
        /// </summary>
        /// <param name="copyItemResponse">The response got from server via CopyItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCopyItemOperation(CopyItemResponseType copyItemResponse, bool isSchemaValidated)
        {
            // If the validation event returns any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R202");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R202            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                202,
                @"[In CopyItem] The following is the WSDL port type specification of the CopyItem operation. <wsdl:operation name=""CopyItem"">
    <wsdl:input message=""tns:CopyItemSoapIn"" />
    <wsdl:output message=""tns:CopyItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R166");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R166            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                166,
                @"[In CopyItem] The following is the WSDL binding specification of the CopyItem operation. <wsdl:operation name=""CopyItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CopyItem"" />
    <wsdl:input>
        <soap:header message=""tns:CopyItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:CopyItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:CopyItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""CopyItemResult"" use=""literal"" />
        <soap:header message=""tns:CopyItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            ItemInfoResponseMessageType item = copyItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            if (item.Items.Items[0] is MessageType)
            {
                MessageType messageItem = item.Items.Items[0] as MessageType;
                this.VerifyMessageType(messageItem, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the CreateItem operation and CreateItemResponseType structure. 
        /// </summary>
        /// <param name="createItemResponse">The response got from server via CreateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyCreateItemOperation(CreateItemResponseType createItemResponse, bool isSchemaValidated)
        {
            // If the validation event returns any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R107");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R107        
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                107,
                @"[In CreateItem] The following is the WSDL port type specification of the CreateItem operation.
<wsdl:operation name=""CreateItem"">
     <wsdl:input message=""tns:CreateItemSoapIn"" />
     <wsdl:output message=""tns:CreateItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R109");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R109         
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                109,
                @"[In CreateItem] The following is the WSDL binding specification of the CreateItem operation. <wsdl:operation name=""CreateItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem"" />
    <wsdl:input>
        <soap:header message=""tns:CreateItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:CreateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:CreateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:header message=""tns:CreateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""CreateItemResult"" use=""literal"" />
        <soap:header message=""tns:CreateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            ItemInfoResponseMessageType item = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            if (item.Items != null && item.Items.Items != null)
            {
                if (item.Items.Items[0] as MessageType != null)
                {
                    MessageType messageItem = item.Items.Items[0] as MessageType;
                    this.VerifyMessageType(messageItem, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the DeleteItem operation. 
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyDeleteItemOperation(bool isSchemaValidated)
        {
            // If the validation event returns any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R148");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R148        
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                148,
                @"[In DeleteItem] The following is the WSDL port type specification of the DeleteItem operation. <wsdl:operation name=""DeleteItem"">
    <wsdl:input message=""tns:DeleteItemSoapIn"" />
    <wsdl:output message=""tns:DeleteItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R150");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R150      
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                150,
                @"[In DeleteItem] The following is the WSDL binding specification of the DeleteItem operation. <wsdl:operation name=""DeleteItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/DeleteItem"" />
    <wsdl:input>
        <soap:header message=""tns:DeleteItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:DeleteItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:DeleteItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""DeleteItemResult"" use=""literal"" />
        <soap:header message=""tns:DeleteItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the GetItem operation and GetItemResponseType structure. 
        /// </summary>
        /// <param name="getItemResponse">The response got from server via GetItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyGetItemOperation(GetItemResponseType getItemResponse, bool isSchemaValidated)
        {
            // If the validation event returns any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R116");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R116           
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                116,
                @"[In GetItem] The following is the WSDL port type specification of the GetItem operation.<wsdl:operation name=""GetItem"">
    <wsdl:input message=""tns:GetItemSoapIn"" />
    <wsdl:output message=""tns:GetItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R118");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R118       
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                118,
                @"[In GetItem] The following is the WSDL binding specification of the GetItem operation. <wsdl:operation name=""GetItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetItem"" />
    <wsdl:input>
        <soap:header message=""tns:GetItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:GetItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:GetItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:header message=""tns:GetItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""GetItemResult"" use=""literal"" />
        <soap:header message=""tns:GetItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            ItemInfoResponseMessageType item = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            if (item.ResponseClass == ResponseClassType.Success)
            {
                if (item.Items.Items[0] is MessageType)
                {
                    MessageType messageItem = item.Items.Items[0] as MessageType;
                    this.VerifyMessageType(messageItem, isSchemaValidated);
                }
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the MoveItem operation and MoveItemResponseType structure. 
        /// </summary>
        /// <param name="moveItemResponse">The response got from server via MoveItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMoveItemOperation(MoveItemResponseType moveItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R157");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R157            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                157,
                @"[In MoveItem] The following is the WSDL port type specification of the MoveItem operation. <wsdl:operation name=""MoveItem"">
    <wsdl:input message=""tns:MoveItemSoapIn"" />
    <wsdl:output message=""tns:MoveItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R159");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R159     
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                159,
                @"[In MoveItem] The following is the WSDL binding specification of the MoveItem operation. <wsdl:operation name=""MoveItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem"" />
    <wsdl:input>
        <soap:header message=""tns:MoveItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:MoveItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:MoveItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""MoveItemResult"" use=""literal"" />
        <soap:header message=""tns:MoveItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            ItemInfoResponseMessageType item = moveItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            if (item.Items.Items[0] is MessageType)
            {
                MessageType messageItem = item.Items.Items[0] as MessageType;
                this.VerifyMessageType(messageItem, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the UpdateItem operation and UpdateItemResponseType structure. 
        /// </summary>
        /// <param name="updateItemResponse">The response got from server via UpdateItem operation.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyUpdateItemOperation(UpdateItemResponseType updateItemResponse, bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R137");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R137            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                137,
                @"[In UpdateItem] The following is the WSDL port type specification of the operation. <wsdl:operation name=""UpdateItem"">
    <wsdl:input message=""tns:UpdateItemSoapIn"" />
    <wsdl:output message=""tns:UpdateItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R139");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R139            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                139,
                @"[In UpdateItem] The following is the WSDL binding specification of the UpdateItem operation. <wsdl:operation name=""UpdateItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem"" />
    <wsdl:input>
        <soap:header message=""tns:UpdateItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:UpdateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:UpdateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:header message=""tns:UpdateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""UpdateItemResult"" use=""literal"" />
        <soap:header message=""tns:UpdateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");

            ItemInfoResponseMessageType item = updateItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            if (item.Items.Items[0] is MessageType)
            {
                MessageType messageItem = item.Items.Items[0] as MessageType;
                this.VerifyMessageType(messageItem, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify the WSDL port type specifications for the SendItem operation. 
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifySendItemOperation(bool isSchemaValidated)
        {
            // If the validation event return any error or warning, the schema validation is false, 
            // which indicates the schema is not matched with the expected result. 
            // So this requirement can be verified when the isSchemaValidation is true.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R173");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R173            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                173,
                @"[In SendItem] The following is the WSDL port type specification of the SendItem operation. <wsdl:operation name=""SendItem "">
    <wsdl:input message=""tns:SendItemSoapIn"" />
    <wsdl:output message=""tns:SendItemSoapOut"" />
</wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R175");

            // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R175            
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                175,
                @"[In SendItem] The following is the WSDL binding specification of the SendItem operation. <wsdl:operation name=""SendItem"">
    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/SendItem"" />
    <wsdl:input>
        <soap:header message=""tns:SendItemSoapIn"" part=""Impersonation"" use=""literal""/>
        <soap:header message=""tns:SendItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
        <soap:header message=""tns:SendItemSoapIn"" part=""RequestVersion"" use=""literal""/>
        <soap:body parts=""request"" use=""literal"" />
    </wsdl:input>
    <wsdl:output>
        <soap:body parts=""SendItemResult"" use=""literal"" />
        <soap:header message=""tns:SendItemSoapOut"" part=""ServerVersion"" use=""literal""/>
    </wsdl:output>
</wsdl:operation>");
        }

        /// <summary>
        /// Verify the requirements about MessageType properties.
        /// </summary>
        /// <param name="messageItem">An instance of MessageType type.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMessageType(MessageType messageItem, bool isSchemaValidated)
        {
            if (messageItem.Sender != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R25");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R25
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    25,
                    @"[In t:MessageType Complex Type] The Sender element is t:SingleRecipientType ([MS-OXWSCDATA] section 2.2.4.71) type.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1292");

                // Verify MS-OXWSMSG requirement: MS-OXWSCDATA_R1292   
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    1292,
                    @"[In t:SingleRecipientType Complex Type] The type [SingleRecipientType] is defined as follow:
 <xs:complexType name=""SingleRecipientType"">
  <xs:choice>
    <xs:element name=""Mailbox""
      type=""t:EmailAddressType""
     />
  </xs:choice>
</xs:complexType>");
            }

            if (messageItem.ToRecipients != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R29");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R29
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    29,
                    @"[In t:MessageType Complex Type] ToRecipients element is t:ArrayOfRecipientsType ([MS-OXWSCDATA] section 2.2.4.11) type.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1033");

                // Verify MS-OXWSMSG requirement: MS-OXWSCDATA_R1033 
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    1033,
                    @"[In t:ArrayOfRecipientsType Complex Type] The type [ArrayOfRecipientsType] is defined as follow:
 <xs:complexType name=""ArrayOfRecipientsType"">
  <xs:choice
    maxOccurs=""unbounded""
    minOccurs=""0""
  >
    <xs:element name=""Mailbox""
      type=""t:EmailAddressType""
     />
  </xs:choice>
</xs:complexType>");
            }

            if (messageItem.CcRecipients != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R184");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R184
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    184,
                    @"[In t:MessageType Complex Type] CcRecipients element is t:ArrayOfRecipientsType type.");
            }

            if (messageItem.BccRecipients != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R185");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R185
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    185,
                    @"[In t:MessageType Complex Type] BccRecipients element is t:ArrayOfRecipientsType type.");
            }

            if (messageItem.IsReadReceiptRequestedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R186");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R186
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    186,
                    @"[In t:MessageType Complex Type] IsReadReceiptRequested element is xs:boolean ([XMLSCHEMA2] sec 3.2.2) type.");
            }

            if (messageItem.IsDeliveryReceiptRequestedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R187");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R187
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    187,
                    @"[In t:MessageType Complex Type] IsDeliveryReceiptRequested element is xs:boolean type.");
            }

            if (messageItem.ConversationIndex != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R188");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R188
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    188,
                    @"[In t:MessageType Complex Type] ConversationIndex element is xs:base64Binary ([XMLSCHEMA2] sec 3.2.16) type.");
            }

            if (messageItem.ConversationTopic != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R189");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R189
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    189,
                    @"[In t:MessageType Complex Type] ConversationTopic element is xs:string ([XMLSCHEMA2] sec 3.2.1) type.");
            }

            if (messageItem.ApprovalRequestData != null && Common.IsRequirementEnabled(182002, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73002");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R73002
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    73002,
                    @"[In t:MessageType Complex Type] The type of child element ApprovalRequestData is t:ApprovalRequestDataType ([MS-OXWSMTGS] section 2.2.4.3).");
            }

            if (messageItem.VotingInformation != null && Common.IsRequirementEnabled(182004, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73004");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R73004
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    73004,
                    @"[In t:MessageType Complex Type] The type of child element VotingInformation is t:VotingInformationType ([MS-OXWSMTGS] section 2.2.4.39)."); 
            }

            if (messageItem.ReminderMessageData != null && Common.IsRequirementEnabled(182006, Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R73006");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R73006
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    73006,
                    @"[In t:MessageType Complex Type] The type of child element ReminderMessageData is t:ReminderMessageDataType ( [MS-OXWSMTGS] section 2.2.4.33).");  
            }
            
            if (messageItem.From != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R190");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R190
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    190,
                    @"[In t:MessageType Complex Type] From element is t:SingleRecipientType type.");
            }

            if (messageItem.InternetMessageId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R191");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R191
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    191,
                    @"[In t:MessageType Complex Type] InternetMessageId element is xs:string type.");
            }

            if (messageItem.IsReadSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R192");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R192
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    192,
                    @"[In t:MessageType Complex Type] IsRead element is xs:boolean type.");
            }

            if (messageItem.IsResponseRequestedSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R193");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R193
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    193,
                    @"[In t:MessageType Complex Type] IsResponseRequested element is xs:boolean type.");
            }

            if (messageItem.References != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R194");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R194
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    194,
                    @"[In t:MessageType Complex Type] References element is xs:string type.");
            }

            if (messageItem.ReplyTo != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R195");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R195
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    195,
                    @"[In t:MessageType Complex Type] ReplyTo element is t:ArrayOfRecipientsType type.");
            }

            if (messageItem != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R22");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R22
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    22,
                    @"[In t:MessageType Complex Type] The MessageType complex type extends the ItemType complex type ([MS-OXWSCORE] section 2.2.4.8).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSMSG_R23");

                // Verify MS-OXWSMSG requirement: MS-OXWSMSG_R23
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    23,
                    @"[In t:MessageType Complex Type] The MessageType schema is <xs:complexType name=""MessageType"">
  <xs:complexContent>
    <xs:extension
      base=""t:ItemType""
    >
      <xs:sequence>
        <xs:element name=""Sender""
          type=""t:SingleRecipientType""
          minOccurs=""0""
         />
        <xs:element name=""ToRecipients""
          type=""t:ArrayOfRecipientsType""
          minOccurs=""0""
         />
        <xs:element name=""CcRecipients""
          type=""t:ArrayOfRecipientsType""
          minOccurs=""0""
         />
        <xs:element name=""BccRecipients""
          type=""t:ArrayOfRecipientsType""
          minOccurs=""0""
         />
        <xs:element name=""IsReadReceiptRequested""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""IsDeliveryReceiptRequested""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""ConversationIndex""
          type=""xs:base64Binary""
          minOccurs=""0""
         />
        <xs:element name=""ConversationTopic""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""From""
          type=""t:SingleRecipientType""
          minOccurs=""0""
         />
        <xs:element name=""InternetMessageId""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""IsRead""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""IsResponseRequested""
          type=""xs:boolean""
          minOccurs=""0""
         />
        <xs:element name=""References""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""ReplyTo""
          type=""t:ArrayOfRecipientsType""
          minOccurs=""0""
         />
        <xs:element name=""ReceivedBy""
          type=""t:SingleRecipientType""
          minOccurs=""0""
         />
        <xs:element name=""ReceivedRepresenting""
          type=""t:SingleRecipientType""
          minOccurs=""0""
         />
        <xs:element name=""ApprovalRequestData""
	          type=""t:ApprovalRequestDataType"" 
	          minOccurs=""0""
	     />
	     <xs:element name=""VotingInformation""
	          type=""t:VotingInformationType"" 
	          minOccurs=""0""
	     />
	     <xs:element name=""ReminderMessageData"" 
	          type=""t:ReminderMessageDataType"" 
	          minOccurs=""0""
	     />
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");
            }

            if (messageItem.Sender != null && messageItem.Sender.Item != null)
            {
                this.VerifyEmailAddressType(messageItem.Sender.Item, isSchemaValidated);
            }

            if (messageItem.From != null && messageItem.From.Item != null)
            {
                this.VerifyEmailAddressType(messageItem.From.Item, isSchemaValidated);
            }

            if (messageItem.ReceivedBy != null && messageItem.ReceivedBy.Item != null)
            {
                this.VerifyEmailAddressType(messageItem.ReceivedBy.Item, isSchemaValidated);
            }

            if (messageItem.ReceivedRepresenting != null && messageItem.ReceivedRepresenting.Item != null)
            {
                this.VerifyEmailAddressType(messageItem.ReceivedRepresenting.Item, isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify EmailAddressType type.
        /// </summary>
        /// <param name="emailAddress">An instance of EmailAddressType type.</param>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyEmailAddressType(EmailAddressType emailAddress, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1147");

            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1147,
                @"[In t:EmailAddessType Complex Type] The type [EmailAddressType] is defined as follow:
<xs:complexType name=""EmailAddressType"">
  <xs:complexContent>
    <xs:extension
      base=""t:BaseEmailAddressType""
    >
      <xs:sequence>
        <xs:element name=""Name""
          type=""xs:string""
          minOccurs=""0""
         />
        <xs:element name=""EmailAddress""
          type=""t:NonEmptyStringType""
          minOccurs=""0""
         />
        <xs:element name=""RoutingType""
          type=""t:NonEmptyStringType""
          minOccurs=""0""
         />
        <xs:element name=""MailboxType""
          type=""t:MailboxTypeType""
          minOccurs=""0""
         />
        <xs:element name=""ItemId""
          type=""t:ItemIdType""
          minOccurs=""0""
         />
        <xs:element name=""OriginalDisplayName"" 
          type=""xs:string"" 
          minOccurs=""0""/>
      </xs:sequence>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");

            if (!string.IsNullOrEmpty(emailAddress.EmailAddress))
            {
                this.VerifyNonEmptyStringType(isSchemaValidated);
            }

            if (emailAddress.MailboxTypeSpecified)
            {
                this.VerifyMailboxTypeType(isSchemaValidated);
            }

            if (!string.IsNullOrEmpty(emailAddress.RoutingType))
            {
                this.VerifyNonEmptyStringType(isSchemaValidated);
            }
        }

        /// <summary>
        /// Verify NonEmptyStringType type.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyNonEmptyStringType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R189");

            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                189,
                @"[In t:NonEmptyStringType Simple Type] The type [NonEmptyStringType] is defined as follow:
<xs:simpleType name=""NonEmptyStringType"">
    <xs:restriction base=""xs:string"">
        <xs:minLength value=""1""/>
    </xs:restriction>
</xs:simpleType>");
        }

        /// <summary>
        /// Verify MailboxTypeType type.
        /// </summary>
        /// <param name="isSchemaValidated">The result of schema validation, true means valid.</param>
        private void VerifyMailboxTypeType(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R163");

            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                163,
                @"[In t:MailboxTypeType Simple Type] The type [MailboxTypeType] is defined as follow:
<xs:simpleType name=""MailboxTypeType"">
    <xs:restriction base=""xs:string"">
        <xs:enumeration value=""Unknown""/>
        <xs:enumeration value=""OneOff""/>
        <xs:enumeration value=""Contact""/>
        <xs:enumeration value=""Mailbox""/>
        <xs:enumeration value=""PrivateDL""/>
        <xs:enumeration value=""PublicDL""/>
        <xs:enumeration value=""PublicFolder""/>
    </xs:restriction>
</xs:simpleType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1428001");

            Site.CaptureRequirementIfIsTrue(
                Common.GetConfigurationPropertyValue("SutVersion", this.Site).Equals("ExchangeServer2007") == false,
                "MS-OXWSCDATA",
                1428001,
                @"[In Appendix C: Product Behavior]<53> Section 2.2.4.31:  Exchange 2010 and above return the MailboxType element in the GetItem operation.");
        }
    }
}