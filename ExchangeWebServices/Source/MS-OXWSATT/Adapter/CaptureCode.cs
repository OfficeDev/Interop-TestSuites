namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSATT.
    /// </summary>
    public partial class MS_OXWSATTAdapter : ManagedAdapterBase, IMS_OXWSATTAdapter
    {
        /// <summary>
        /// The capture code for requirements of CreateAttachment operation.
        /// </summary>
        /// <param name="createAttachmentResponse">CreateAttachmentResponseType createAttachmentResponse</param>
        /// <param name="isSchemaValidated">Indicate whether the schema is verified</param>
        private void VerifyCreateAttachmentResponse(CreateAttachmentResponseType createAttachmentResponse, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R372");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R372
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                372,
                @"[In tns:CreateAttachmentSoapOut Message][The element of CreateAttachmentResult part is] tns:CreateAttachmentResponse (section 3.1.4.1.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R553");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R553
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                553,
                @"[In tns:CreateAttachmentSoapOut Message][The element of ServerVersion part is] t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R542");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R542
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                542,
                @"[In tns:CreateAttachmentSoapOut Message][The ServerVersion part] Specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R543");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R543
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                543,
                @"[In Elements] [Element name] CreateAttachmentResponse Specifies the response body content from a request to create an attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R545");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R545
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                545,
                @"[In Complex Types] [Complex type name] CreateAttachmentResponseType [Description] Specifies a response message for the CreateAttachment operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R535");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R535
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                535,
                @"[In CreateAttachment Operation] The following is the WSDL port type specification of the  CreateAttachment operation.
                    <wsdl:operation name=""CreateAttachment"">
                            <wsdl:input message=""tns:CreateAttachmentSoapIn"" />
                            <wsdl:output message=""tns:CreateAttachmentSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R320");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R320
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                320,
                @"[In CreateAttachment Operation] The following is the WSDL binding specification of the CreateAttachment operation.
                    <wsdl:operation name=""CreateAttachment"">
                        <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateAttachment"" />
                        <wsdl:input>
                            <soap:header message=""tns:CreateAttachmentSoapIn"" part=""Impersonation"" use=""literal""/>
                            <soap:header message=""tns:CreateAttachmentSoapIn"" part=""MailboxCulture"" use=""literal""/>
                            <soap:header message=""tns:CreateAttachmentSoapIn"" part=""RequestVersion"" use=""literal""/>
                            <soap:header message=""tns:CreateAttachmentSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                            <soap:body parts=""request"" use=""literal"" />
                        </wsdl:input>
                        <wsdl:output>
                            <soap:body parts=""CreateAttachmentResult"" use=""literal"" />
                            <soap:header message=""tns:CreateAttachmentSoapOut"" part=""ServerVersion"" use=""literal""/>
                        </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R539");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R539
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                539,
                @"[In tns:CreateAttachmentSoapIn Message][The RequestVersion part] Specifies a SOAP header that identifies the schema version for the CreateAttachment operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R537");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R537
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                537,
                @"[In Messages][The CreateAttachmentSoapOut message] Specifies the SOAP message that is returned by the server in response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R165");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R165
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                165,
                @"[In tns:CreateAttachmentSoapOut Message] The CreateAttachmentSoapOut WSDL message specifies the server response to the CreateAttachment operation request to create an attachment.
                    <wsdl:message name=""CreateAttachmentSoapOut"">
                        <wsdl:part name=""CreateAttachmentResult"" element=""tns:CreateAttachmentResponse"" />
                        <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R182");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R182
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                182,
                @"[In m:CreateAttachmentResponseType Complex Type][The CreateAttachmentResponseType is defined as follow:]
                    <xs:complexType name=""CreateAttachmentResponseType"">
                      <xs:complexContent>
                        <xs:extension
                    base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            AttachmentIdType attachmentId = (createAttachmentResponse.ResponseMessages.Items[0] as AttachmentInfoResponseMessageType).Attachments[0].AttachmentId;

            // AttachmentIdType is optional in AttachmentType
            if (attachmentId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R333");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R333
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    333,
                    @"[In t:AttachmentIdType Complex Type][The type of RootItemId attribute is] xs:string ([XMLSCHEMA2]).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R439");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R439
                // Validate the length of RootItemId is no more than 512 bytes after base64 decoding.
                Site.Log.Add(LogEntryKind.Debug, "Length of root item id is:{0}", attachmentId.RootItemId.Length);
                bool isVerifyR439 = attachmentId.RootItemId.Length <= 512;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR439,
                    439,
                    @"[In t:AttachmentIdType Complex Type][In RootItemId] The length of RootItemId attribute is no more than 512 bytes after base64 decoding.");

                if (attachmentId.RootItemChangeKey != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R334");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R334
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        334,
                        @"[In t:AttachmentIdType Complex Type][The type of RootItemChangeKey attribute is] xs:string.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R440");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R440
                    // Validate the length of RootItemChangeKey is no more than 512 bytes after base64 decoding.
                    Site.Log.Add(LogEntryKind.Debug, "Length of root item change key is:{0}", attachmentId.RootItemChangeKey.Length);
                    bool isVerifyR440 = attachmentId.RootItemChangeKey.Length <= 512;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR440,
                        440,
                        @"[In t:AttachmentIdType Complex Type][In RootItemChangeKey] The length of RootItemChangeKey attribute is no more than 512 bytes after base64 decoding.");
                }
            }
        }

        /// <summary>
        /// The capture code for requirements of DeleteAttachment operation.
        /// </summary>
        /// <param name="deleteAttachmentResponse">DeleteAttachmentResponseType deleteAttachmentResponse</param>
        /// <param name="isSchemaValidated">Indicate whether the schema is verified</param>
        private void VerifyDeleteAttachmentResponse(DeleteAttachmentResponseType deleteAttachmentResponse, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R382");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R382
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                382,
                @"[In tns:DeleteAttachmentSoapOut Message][The element of DeleteAttachmentResult part is] tns:DeleteAttachmentResponse (section 3.1.4.2.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R464");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R464
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                464,
                @"[In tns:DeleteAttachmentSoapOut Message][The DeleteAttachmentResult part] Specifies the SOAP body of the response to a DeleteAttachment operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R383");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R383
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                383,
                @"[In tns:DeleteAttachmentSoapOut Message][The element of ServerVersion part is] t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R465");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R465
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                465,
                @"[In tns:DeleteAttachmentSoapOut Message][The ServerVersion part] Specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R507");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R507
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                507,
                @"[In Elements] [Element name] [DeleteAttachmentResponse] Specifies the response body content from a request to delete an attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R323");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R323
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                323,
                @"[In DeleteAttachment Operation] The following is the WSDL port type specification of the operation.
                    <wsdl:operation name=""DeleteAttachment"">
                        <wsdl:input message=""tns:DeleteAttachmentSoapIn"" />
                        <wsdl:output message=""tns:DeleteAttachmentSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R324");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R324
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                324,
                @"[In DeleteAttachment Operation] The following is the WSDL binding specification of the operation.
                    <wsdl:operation name=""DeleteAttachment"">
                        <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/DeleteAttachment"" />
                        <wsdl:input>
                            <soap:header message=""tns:DeleteAttachmentSoapIn"" part=""Impersonation"" use=""literal""/>
                            <soap:header message=""tns:DeleteAttachmentSoapIn"" part=""MailboxCulture"" use=""literal""/>
                            <soap:header message=""tns:DeleteAttachmentSoapIn"" part=""RequestVersion"" use=""literal""/>
                            <soap:body parts=""request"" use=""literal"" />
                        </wsdl:input>
                        <wsdl:output>
                            <soap:body parts=""DeleteAttachmentResult"" use=""literal"" />
                            <soap:header message=""tns:DeleteAttachmentSoapOut"" part=""ServerVersion"" use=""literal""/>
                        </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R462");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R462
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                462,
                @"[In tns:DeleteAttachmentSoapIn Message][The RequestVersion part] Specifies a SOAP header that identifies the schema version for the DeleteAttachment operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R207");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R207
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                207,
                @"[In Messages][The DeleteAttachmentSoapOut  message] Specifies the SOAP message that is returned by the server in response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R222");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R222
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                222,
                @"[In tns:DeleteAttachmentSoapOut Message] The DeleteAttachmentSoapOut WSDL message specifies the server response to the DeleteAttachment operation request to delete an attachment.
                     <wsdl:message name=""DeleteAttachmentSoapOut"">
                        <wsdl:part name=""DeleteAttachmentResult"" element=""tns:DeleteAttachmentResponse"" />
                        <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R326");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R326
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                326,
                @"[In DeleteAttachmentResponse Element] The DeleteAttachmentResponse element specifies the response message for a DeleteAttachment operation.
                     <xs:element name=""DeleteAttachmentResponse""
                      type=""m:DeleteAttachmentResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R510");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R510
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                510,
                @"[In Complex Types] [Complex type name]DeleteAttachmentResponseType Specifies a response message for the DeleteAttachment operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R246");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R246
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                246,
                @"[In m:DeleteAttachmentResponseType Complex Type][The DeleteAttachmentResponseType is defined as follow:]
                    <xs:complexType name=""DeleteAttachmentResponseType"">
                      <xs:complexContent>
                        <xs:extension
                        base=""m:BaseResponseMessageType""
                        />
                      </xs:complexContent>
                    </xs:complexType>");

            foreach (DeleteAttachmentResponseMessageType message in deleteAttachmentResponse.ResponseMessages.Items)
            {
                this.VerifyDeleteAttachmentResponseMessageType(message, isSchemaValidated);
            }
        }

        /// <summary>
        /// Captures DeleteAttachmentResponseMessageType related requirements.
        /// </summary>
        /// <param name="deleteAttachmentResponseMessage">An DeleteAttachmentResponseMessageType instance.</param>
        /// <param name="isSchemaValidated">A Boolean value indicates the schema validation result.</param>
        private void VerifyDeleteAttachmentResponseMessageType(DeleteAttachmentResponseMessageType deleteAttachmentResponseMessage, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R387");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R387
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                387,
                @"[In m:DeleteAttachmentResponseMessageType Complex Type][The type of RootItemId element is] t:RootItemIdType (section 2.2.4.9).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R240");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R240
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                240,
                @"[In m:DeleteAttachmentResponseMessageType Complex Type][The DeleteAttachmentResponseMessageType is defined as follow:]
                    <xs:complexType name=""DeleteAttachmentResponseMessageType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:ResponseMessageType""
                        >
                          <xs:sequence>
                            <xs:element name=""RootItemId""
                              type=""t:RootItemIdType""
                              minOccurs=""0""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            RootItemIdType rootItemId = deleteAttachmentResponseMessage.RootItemId;

            // RootItemId is optional in DeleteAttachmentResponse
            if (rootItemId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R482");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R482
                Site.Log.Add(LogEntryKind.Debug, "Length of root item id is:{0}", rootItemId.RootItemId.Length);
                bool isVerifyR482 = rootItemId.RootItemId.Length <= 512;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR482,
                    482,
                    @"[In t:RootItemIdType Complex Type][In RootItemId] The length of RootItemId is no more than 512 bytes after base64 decoding.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R556");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R556
                // Validate the RootItemChangeKey is not null and its length is no more than 512 bytes after base64 decoding.
                Site.Log.Add(LogEntryKind.Debug, "Length of root item change key is:{0}", rootItemId.RootItemChangeKey.Length);
                bool isVerifyR556 = rootItemId.RootItemChangeKey.Length <= 512;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR556,
                    556,
                    @"[In t:RootItemIdType Complex Type][In RootItemChangeKey] The length of RootItemChangeKey attribute is no more than 512 bytes after base64 decoding.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R533");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R533
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    533,
                    @"[In Message Processing Events and Sequencing Rules][The DeleteAttachment operation] Deletes file and item attachments from an existing item in the server store.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R442");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R442
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    442,
                    @"[In t:RootItemIdType Complex Type][The type of RootItemId attribute is] xs:string ([XMLSCHEMA2]).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R445");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R445
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    445,
                    @"[In t:RootItemIdType Complex Type][The type of RootItemChangeKey attribute is] xs:string.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R113");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R113
                // The type of RootItemId is RootItemIdType, as verified in MS-OXWSATT_R387.
                // validateSchema has already validate the DeleteAttachmentResponse is not null and matches the schema.
                // And the elements and their types have been verified in MS-OXWSATT_R442 and MS-OXWSATT_R445.
                // Thus this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    113,
                    @"[In t:RootItemIdType Complex Type][The RootItemIdType is defined as follow:]
                    <xs:complexType name=""RootItemIdType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:BaseItemIdType""
                        >
                          <xs:attribute name=""RootItemId""
                            type=""xs:string""
                            use=""required""
                           />
                          <xs:attribute name=""RootItemChangeKey""
                            type=""xs:string""
                            use=""required""
                           />
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");
            }
        }

        /// <summary>
        /// The capture code for requirements of GetAttachment operation.
        /// </summary>
        /// <param name="getAttachmentResponse">GetAttachmentResponseType getAttachmentResponse.</param>
        /// <param name="isSchemaValidated">Indicate whether the schema is verified.</param>
        private void VerifyGetAttachmentResponse(GetAttachmentResponseType getAttachmentResponse, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R514");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R514
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                514,
                @"[In Messages] [Message name] GetAttachmentSoapOut Specifies the SOAP message that is returned by the server in response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R534");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R534
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                534,
                @"[In Message Processing Events and Sequencing Rules][The GetAttachment operation] Retrieves existing attachments on items in the server store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R471");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R471
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                471,
                @"[In tns:GetAttachmentSoapIn Message] [The RequestVersion part] Specifies a SOAP header that identifies the schema version for the GetAttachment operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R396");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R396
            // In SchemaValidation.cs, the nodesForSoapHeader, which is the element of SOAP header, have been validated by XmlValidater(schemaListCotent, header.InnerXml).
            // Thus R396 can be captured according validate the schema.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                396,
                @"[In tns:GetAttachmentSoapOut Message][The element of ServerVersion part is] t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R395");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R395
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                395,
                @"[In tns:GetAttachmentSoapOut Message] [The element of GetAttachmentResult part is] tns:GetAttachmentResponse (section 3.1.4.3.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R473");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R473
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                473,
                @"[In tns:GetAttachmentSoapOut Message][The GetAttachmentResult part] Specifies the SOAP body of the response to a GetAttachment operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R474");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R474
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                474,
                @"[In tns:GetAttachmentSoapOut Message][The ServerVersion part] Specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R517");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R517
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                517,
                @"[In Elements] [Element name] GetAttachmentResponse Specifies the response body content from a request to get an attachment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R519");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R519
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                519,
                @"[In Complex Types] [Complex type name]GetAttachmentResponseType Specifies a response message for the GetAttachment operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R327");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R327
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                327,
                @"[In GetAttachment Operation] The following is the WSDL port type specification of the operation.
                    <wsdl:operation name=""GetAttachment"">
                        <wsdl:input message=""tns:GetAttachmentSoapIn"" />
                        <wsdl:output message=""tns:GetAttachmentSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R328");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R328
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                328,
                @"[In GetAttachment Operation] The following is the WSDL binding specification of the operation.
                    <wsdl:operation name=""GetAttachment"">
                        <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetAttachment"" />
                        <wsdl:input>
                            <soap:header message=""tns:GetAttachmentSoapIn"" part=""Impersonation"" use=""literal""/>
                            <soap:header message=""tns:GetAttachmentSoapIn"" part=""MailboxCulture"" use=""literal""/>
                            <soap:header message=""tns:GetAttachmentSoapIn"" part=""RequestVersion"" use=""literal""/>
                            <soap:header message=""tns:GetAttachmentSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                            <soap:body parts=""request"" use=""literal"" />
                        </wsdl:input>
                        <wsdl:output>
                            <soap:body parts=""GetAttachmentResult"" use=""literal"" />
                            <soap:header message=""tns:GetAttachmentSoapOut"" part=""ServerVersion"" use=""literal""/>
                        </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R280");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R280
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                280,
                @"[In tns:GetAttachmentSoapOut Message] The GetAttachmentSoapOut WSDL message specifies the server response to the GetAttachment operation request to get an attachment.
                    <wsdl:message name=""GetAttachmentSoapOut"">
                        <wsdl:part name=""GetAttachmentResult"" element=""tns:GetAttachmentResponse"" />
                        <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R330");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R330
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                330,
                @"[In GetAttachmentResponse Element] The GetAttachmentResponse element specifies the response message to a GetAttachment operation. <xs:element name=""GetAttachmentResponse""
                      type=""m:GetAttachmentResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R298");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R298
            // The type of GetAttachmentResponse is GetAttachmentResponseType.
            // validateSchema has already validate the GetAttachmentResponse is not null and matches the schema.
            // Thus this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                298,
                @"[In m:GetAttachmentResponseType Complex Type] The GetAttachmentResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).
                     <xs:complexType name=""GetAttachmentResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            AttachmentInfoResponseMessageType attachmentInfo = (AttachmentInfoResponseMessageType)getAttachmentResponse.ResponseMessages.Items[0];
            if (attachmentInfo.Attachments.Length > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R52");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R52
                // Validate the type of Attachment is not an ItemType.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    52,
                    @"[In m:AttachmentInfoResponseMessageType Complex Type][The Attachments element] Represents an array of types based on attachments on the item.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R5577");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R5577
            // Schema is verified in adapter, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                5577,
                @"[In m:AttachmentInfoResponseMessageType Complex Type] The ArrayOfAttachmentsType complex type is used in the response message.");

            if (attachmentInfo.Attachments.Length > 0)
            {
                if (attachmentInfo.Attachments[0].GetType() == typeof(FileAttachmentType))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R332");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R332
                    // Validate the type of fileAttach is FileAttachmentType.
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        332,
                        @"[In t:ArrayOfAttachmentsType Complex Type][The type of FileAttachment element is] t:FileAttachmentType (section 2.2.4.5).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R72");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R72
                    // Thus this requirement can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        72,
                        @"[In t:FileAttachmentType Complex Type] The FileAttachmentType complex type extends the AttachmentType complex type, as specified in section 2.2.4.4.
                    <xs:complexType name=""FileAttachmentType"">
                      <xs:complexContent>
                        <xs:extension name=""FileAttachmentType""
                          base=""t:AttachmentType""
                        >
                          <xs:sequence>
                            <xs:element name=""IsContactPhoto""
                              type=""xs:boolean""
                              minOccurs=""0""
                              maxOccurs=""1""
                             />
                            <xs:element name=""Content""
                              type=""xs:base64Binary""
                              minOccurs=""0""
                              maxOccurs=""1""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");
                    FileAttachmentType fileAttachment = (FileAttachmentType)attachmentInfo.Attachments[0];

                    if (fileAttachment.Content != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R347");

                        // Verify MS-OXWSATT requirement: MS-OXWSATT_R347
                        // Validate the type of Content is byte[].
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            347,
                            @"[In t:FileAttachmentType Complex Type][The type of Content element is] xs:base64Binary ([XMLSCHEMA2]).");
                    }
                }

                AttachmentType attachment = attachmentInfo.Attachments[0];

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R55");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R55
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    55,
                    @"[In t:AttachmentType Complex Type] The AttachmentType complex type represents an attachment.
                    <xs:complexType name=""AttachmentType"">
                        <xs:sequence>
                        <xs:element name=""AttachmentId""
                            type=""t:AttachmentIdType""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""Name""
                            type=""xs:string""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""ContentType""
                            type=""xs:string""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""ContentId""
                            type=""xs:string""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""ContentLocation""
                            type=""xs:string""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""Size""
                            type=""xs:int""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""LastModifiedTime""
                            type=""xs:dateTime""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        <xs:element name=""IsInline""
                            type=""xs:boolean""
                            minOccurs=""0""
                            maxOccurs=""1""
                            />
                        </xs:sequence>
                    </xs:complexType>");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R336");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R336
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    336,
                    @"[In t:AttachmentType Complex Type][The type of AttachmentId element is] t:AttachmentIdType (section 2.2.4.2).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R41");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R41
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    41,
                    @"[In t:AttachmentIdType Complex Type] [the schema of ""AttachmentIdType"" is:]
                    <xs:complexType name=""AttachmentIdType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:RequestAttachmentIdType""
                        >
                          <xs:attribute name=""RootItemId""
                            type=""xs:string""
                            use=""optional""
                           />
                          <xs:attribute name=""RootItemChangeKey""
                            type=""xs:string""
                            use=""optional""
                           />
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

                if (attachment.Name != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R337");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R337
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        337,
                        @"[In t:AttachmentType Complex Type][The type of Name element is] xs:string([XMLSCHEMA2]).");
                }

                if (attachment.ContentType != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R338");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R338
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        338,
                        @"[In t:AttachmentType Complex Type][The type of ContentType element is] xs:string.");
                }

                if (attachment.ContentId != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R339");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R339
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        339,
                        @"[In t:AttachmentType Complex Type][The type of ContentId element is] xs:string.");
                }

                if (attachment.ContentLocation != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R340");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R340
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        340,
                        @"[In t:AttachmentType Complex Type][The type of ContentLocation element is] xs:string.");
                }

                if (attachment.LastModifiedTime != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R342");

                    // Verify MS-OXWSATT requirement: MS-OXWSATT_R342
                    Site.CaptureRequirementIfIsTrue(
                        isSchemaValidated,
                        342,
                        @"[In t:AttachmentType Complex Type][The type of LastModifiedTime element is] xs:dateTime ([XMLSCHEMA2]).");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R343");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R343
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    343,
                    @"[In t:AttachmentType Complex Type][The type of IsInline element is] xs:Boolean ([XMLSCHEMA2]).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R49");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R49
                // The elements and their types have been verified in MS-OXWSATT_R335.
                // Thus this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    49,
                    @"[In m:AttachmentInfoResponseMessageType Complex Type] The AttachmentInfoResponseMessageType complex type extends the ResponseMessageType complex type, ([MS-OXWSCDATA] section 2.2.4.57).
                    <xs:complexType name=""AttachmentInfoResponseMessageType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:ResponseMessageType""
                        >
                          <xs:sequence>
                            <xs:element name=""Attachments""
                              type=""t:ArrayOfAttachmentsType""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R335");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R335
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    335,
                    @"[In m:AttachmentInfoResponseMessageType Complex Type][The type of Attachments element is]  t:ArrayOfAttachmentsType (section 2.2.4.1).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R441");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R441
                // If rspMessage.Attachments[0] is not null, and ArrayOfAttachmentsType is the type of Attachments,
                // that means ArrayOfAttachmentsType is used in response.
                // Thus this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    441,
                    @"[In t:ArrayOfAttachmentsType Complex Type] The ArrayOfAttachmentsType complex type is used in the response message.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R33");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R33
                // validateSchema has already validate the GetAttachmentResponse is not null and matches the schema.
                // The elements and their types have been verified in MS-OXWSATT_R331 and MS-OXWSATT_R332.
                // Thus this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    33,
                    @"[In t:ArrayOfAttachmentsType Complex Type][The ArrayOfAttachmentsType Complex Type is defined as follow:]
                    <xs:complexType name=""ArrayOfAttachmentsType"">
                        <xs:choice
                        minOccurs=""0""
                        maxOccurs=""unbounded""
                        >
                        <xs:element name=""ItemAttachment""
                            type=""t:ItemAttachmentType""
                            />
                        <xs:element name=""FileAttachment""
                            type=""t:FileAttachmentType""
                            />
                        </xs:choice>
                    </xs:complexType>");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R331");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R331
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    331,
                    @"[In t:ArrayOfAttachmentsType Complex Type][The type of ItemAttachment element is] t:ItemAttachmentType (section 2.2.4.6).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R79");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R79
                this.Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    79,
                    @"[In t:ItemAttachmentType Complex Type][The ItemAttachmentType is defined as follow:] <xs:complexType name=""ItemAttachmentType"">
  <xs:complexContent>
    <xs:extension
      base=""t:AttachmentType""
    >
      <xs:choice
        minOccurs=""0""
        maxOccurs=""1""
      >
        <xs:element name=""Item""
          type=""t:ItemType""
         />
        <xs:element name=""Message""
          type=""t:MessageType""
         />
        <xs:element name=""CalendarItem""
          type=""t:CalendarItemType""
         />
        <xs:element name=""Contact""
          type=""t:ContactItemType""
         />
        <xs:element name=""MeetingMessage""
          type=""t:MeetingMessageType""
         />
        <xs:element name=""MeetingRequest""
          type=""t:MeetingRequestMessageType""
         />
        <xs:element name=""MeetingResponse""
          type=""t:MeetingResponseMessageType""
         />
        <xs:element name=""MeetingCancellation""
          type=""t:MeetingCancellationMessageType""
         />
        <xs:element name=""Task""
          type=""t:TaskType""
         />
        <xs:element name=""PostItem""
          type=""t:PostItemType""
         />
      </xs:choice>
    </xs:extension>
  </xs:complexContent>
</xs:complexType>");
            }
        }

        /// <summary>
        /// Verify the ServerVersionInfo structure.
        /// </summary>
        /// <param name="isSchemaValidated">Indicate whether the schema is verified.</param>
        private void VerifyServerVersionInfo(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1339");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1339
            // Because the adapter uses SOAP and HTTPS to communicate with server, if server returns data without exception, this requirement will be captured.
            Site.CaptureRequirementIfIsTrue(
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

        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R3");

            // Verify MS-OXWSATT requirement: MS-OXWSATT_R3
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
               3,
               @"[In Transport] The SOAP version supported is SOAP 1.1.");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type.
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (Common.IsRequirementEnabled(318, this.Site) && transport == TransportProtocol.HTTPS)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSATT_R318");

                // Verify MS-OXWSATT requirement: MS-OXWSATT_R318
                // Because Adapter uses SOAP and HTTPS to communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    318,
                    @"[In Appendix C: Product Behavior]Implementation does use secure communications via HTTPS, as defined in [RFC2818]. (Exchange Server 2007 and above follow this behavior. )");
            }
        }
    }
}