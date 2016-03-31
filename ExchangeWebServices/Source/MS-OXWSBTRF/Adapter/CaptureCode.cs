namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSBTRF.
    /// </summary>
    public partial class MS_OXWSBTRFAdapter
    {
        /// <summary>
        /// Verify the SOAP version
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R1");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R1
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                1,
                @"[In Transport]The SOAP version used for this protocol is SOAP 1.1, as specified in [SOAP1.1].");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (transport == TransportProtocol.HTTPS)
            {
                if (Common.IsRequirementEnabled(6, this.Site))
                {
                    // Add debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R6");

                    // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R6
                    // Because Adapter uses SOAP and HTTPS to communicate with server, if server returned data without exception, this requirement has been captured.
                    Site.CaptureRequirement(
                        6,
                        @"[In Appendix C: Product Behavior] Implementation does use secure communications via HTTPS, as defined in [RFC2818]. (Exchange Server 2010 and above follow this behavior.)");
                }
            }
            else if (transport == TransportProtocol.HTTP)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R2006");

                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R2006
                // Because Adapter uses SOAP and HTTP to communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    2006,
                    @"[In Transport]The protocol [MS-OXWSBTRF] MUST support SOAP over HTTP, as specified in [RFC2616].");
            }
        }

        /// <summary>
        /// Verify the ExportItems operation responses related requirements.
        /// </summary>
        /// <param name="exportItems"> The ExportItemsResponseType object indicates ExportItems operation response.</param>
        /// <param name="isSchemaValidated"> A Boolean value indicates whether the schema has been verified.</param>
        private void VerifyExportItemsResponseType(ExportItemsResponseType exportItems, bool isSchemaValidated)
        {
            // Assert ExportItems operation response is not null.
            Site.Assert.IsNotNull(exportItems, "The response of ExportItems operation should not be null.");

            // If the proxy can communicate with server successfully, then all the WSDL related requirements can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R36001");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R36001
            Site.CaptureRequirement(
                36001,
                @"[In ExportItems]The following is the WSDL port type definition of the operation.
                <wsdl:operation name=""ExportItems"">
                    <wsdl:input message=""tns:ExportItemsSoapIn""/>
                    <wsdl:output message=""tns:ExportItemsSoapOut""/>
                </wsdl:operation>");

            // If the schema validation is successful, then MS-OXWSBTRF_R51, MS-OXWSBTRF_R251, MS-OXWSBTRF_R252, MS-OXWSBTRF_R253 
            // and MS-OXWSBTRF_R254 can be captured.
            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R51");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R51
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                51,
                @"[In tns:ExportItemsSoapOut Message]The ExportItemsSoapOut WSDL message specifies the SOAP message that represents a response that contains exported items.
                <wsdl:message name=""ExportItemsSoapOut"">
                    <wsdl:part name=""ExportItemsResult"" element=""tns:ExportItemsResponse""/>
                    <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                </wsdl:message>");

            // Add debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R251");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R251
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                251,
                @"[In tns:ExportItemsSoapOut Message]The Element/type of ExportItemsResult is tns:ExportItemsResponse (section 3.1.4.1.2.2).");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R252");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R252
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                252,
                @"[In tns:ExportItemsSoapOut Message][The part ExportItemsResult] specifies the SOAP body of the response.");
            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R253");

                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R253
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    253,
                    @"[In tns:ExportItemsSoapOut Message]The type of ServerVersion is ServerVersion ([MS-OXWSCDATA] section 2.2.3.12).");

                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R254");

                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R254
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    254,
                    @"[In tns:ExportItemsSoapOut Message][The part ServerVersion]specifies the SOAP header that identifies the server version for the response.");
            }

            // Verify BaseResponseMessageType related requirements in the protocol document MS-OXWSCDATA.
            this.VerifyBaseResponseMessageType(isSchemaValidated);

            // Verify ExportItemsResponseMessageType related requirements.
            foreach (ExportItemsResponseMessageType exportItem in exportItems.ResponseMessages.Items)
            {
                this.VerifyExportItemsResponseMessageType(exportItem, isSchemaValidated);
            }

            // If the schema validation and the above BaseResponseMessageType validation is successful, then MS-OXWSBTRF_R175 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R75");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R75
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                75,
                @"[In ExportItemsResponseType Complex Type][The schema of ExportItemsResponseType is:]<xs:complexType name=""ExportItemsResponseType""><xs:complexContent><xs:extension
                    base=""m:BaseResponseMessageType""/></xs:complexContent>
                    </xs:complexType>");

            // If the schema validation is successful, then the type of ExportItemsResponse element is ExportItemsResponseType, then MS-OXWSBTRF_R63 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R63");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R63
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                63,
                @"[In ExportItemsResponse Element][The schema of ExportItemsResponse is:]<xs:element name=""ExportItemsResponse""
                type=""m:ExportItemsResponseType""/>");

            // If the response from the ExportItems is not null, then MS-OXWSBTRF_R165 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R165");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R165
            Site.CaptureRequirementIfIsNotNull(
                exportItems,
                165,
                @"[In ExportItemsResponse Element]This element [ExportItemsResponse] MUST be present.");

            // If the schema validation is successful, the server version related requirements can be captured.
            this.VerifyServerVersion(isSchemaValidated);
        }

        /// <summary>
        /// This method is used to verify the UploadItems operation response related requirements.
        /// </summary>
        /// <param name="uploadItems"> Specified UploadItemsResponseType instance</param>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyUploadItemsResponseType(UploadItemsResponseType uploadItems, bool isSchemaValidated)
        {
            // Response after Upload Items operation must not be null.
            Site.Assert.IsNotNull(uploadItems, "The response of UploadItems operation should not be null.");

            // If the proxy can communicate with server successfully, then all the WSDL related requirements can be directly captured.
            // Therefore MS-OXWSBTRF_R87, MS-OXWSBTRF_R86 and MS-OXWSBTRF_R101 are captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R87");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R87
            Site.CaptureRequirement(
                87,
                @"[In UploadItems]The following is the WSDL binding specification of the operation.
                <wsdl:operation name=""UploadItems"">
                    <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/UploadItems"" />
                    <wsdl:input>
                    <soap:header message=""tns:UploadItemsSoapIn"" part=""Impersonation"" use=""literal""/>
                    <soap:header message=""tns:UploadItemsSoapIn"" part=""MailboxCulture"" use=""literal""/>
                    <soap:header message=""tns:UploadItemsSoapIn"" part=""RequestVersion"" use=""literal""/>
                    <soap:body parts=""request"" use=""literal"" />
                    </wsdl:input>
                    <wsdl:output>
                    <soap:body parts=""UploadItemsResult"" use=""literal"" />
                    <soap:header message=""tns:UploadItemsSoapOut"" part=""ServerVersion"" use=""literal""/>
                    </wsdl:output>
                </wsdl:operation>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R86");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R86
            Site.CaptureRequirement(
                86,
                @"[In UploadItems]The following is the WSDL port type definition of the operation.<wsdl:operation name=""UploadItems""><wsdl:input message=""tns:UploadItemsSoapIn""/><wsdl:output message=""tns:UploadItemsSoapOut""/></wsdl:operation>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R101");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R101
            Site.CaptureRequirement(
                101,
                @"[In tns:UploadItemsSoapOut Message]The UploadItemsSoapOut WSDL message specifies the SOAP message that represents a response that contains the results of an attempt to upload items into a mailbox.
                [In tns:UploadItemsSoapOut Message]<wsdl:message name=""UploadItemsSoapOut"">
                  <wsdl:part name=""UploadItemsResult"" element=""tns:UploadItemsResponse""/>
                  <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                </wsdl:message>");

            // If the schema validation is successful, then MS-OXWSBTRF_R247, MS-OXWSBTRF_R248 ,MS-OXWSBTRF_R249 and MS-OXWSBTRF_R250 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R247");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R247
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                247,
                @"[In tns:UploadItemsSoapOut Message]The Element/type of UploadItemsResult is
                UploadItemsResponse (section 3.1.4.2.2.2).");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R248");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R248
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                248,
                @"[In tns:UploadItemsSoapOut Message]The UploadItemsResult part specifies the SOAP body of the response to an UploadItems operation request.");
            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R249");

                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R249
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    249,
                    @"[In tns:UploadItemsSoapOut Message]The Element/type of ServerVersion is ServerVersion ([MS-OXWSCDATA] section 2.2.3.12).");

                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R250");

                // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R250
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    250,
                    @"[In tns:UploadItemsSoapOut Message]The ServerVersion part specifies the SOAP header that identifies the server version for the response.");
            }

            // Verify BaseResponseMessageType related requirements in the protocol document MS-OXWSCDATA.
            this.VerifyBaseResponseMessageType(isSchemaValidated);

            // Verify UploadItemsResponseMessageType related requirements.
            foreach (UploadItemsResponseMessageType uploadItem in uploadItems.ResponseMessages.Items)
            {
                this.VerifyUploadItemsResponseMessageType(uploadItem, isSchemaValidated);
            }

            // If the schema validation is successful and the above BaseResponseMessageType validation is fine, then MS-OXWSBTRF_R131 will be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R131");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R131
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                131,
                @"[In m:UploadItemsResponseType Complex Type][The schema of UploadItemsResponseType is:]<xs:complexType name=""UploadItemsResponseType"">
                  <xs:complexContent>
                    <xs:extension base=""m:BaseResponseMessageType""/>
                  </xs:complexContent>
                </xs:complexType>");

            // If the schema validation is successful, then the UpdateItemsResponse element's type is UpdateItemsResponseType, requirement MS-OXWSBTRF_R114 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R114");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R114
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                114,
                @"[In UploadItemsResponse Element] [The schema of UploadItemsResponse is:]
                <xs:element name=""UploadItemsResponse""
                type=""m:UploadItemsResponseType""/>");

            // If the response from the UploadItems is not null, then MS-OXWSBTRF_R246 MS-OXWSBTRF_R11301 and  will be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R246");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R246
            Site.CaptureRequirementIfIsNotNull(
                uploadItems,
                246,
                @"[In Elements]This element [UploadItemsResponse] MUST be present.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R11301");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R11301
            Site.CaptureRequirementIfIsNotNull(
                uploadItems,
                11301,
                @"[In UploadItemsResponse Element] This element [UploadItemsResponse Element] MUST be present.");

            // If the schema validation successful, the server version related requirements can be verified.
            this.VerifyServerVersion(isSchemaValidated);
        }

        /// <summary>
        /// This method is used to verify the ExportItemsResponseMessageType related requirements.
        /// </summary>
        /// <param name="exportItemsMessage">Specified ExportItemsResponseMessageType instance.</param>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyExportItemsResponseMessageType(ExportItemsResponseMessageType exportItemsMessage, bool isSchemaValidated)
        {
            // Verify the base type ResponseMessageType related requirements.
            this.VerifyResponseMessageType(exportItemsMessage as ResponseMessageType, isSchemaValidated);

            // If the schema validation and the above base type verification are successful, then MS-OXWSBTRF_R70, MS-OXWSBTRF_R168 and MS-OXWSBTRF_R171 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R168");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R168
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                168,
                @"[In m:ExportItemsResponseMessageType Complex Type]The type of ItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R70");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R70
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                70,
                @"[In m:ExportItemsResponseMessageType Complex Type][The schema of ExportItemsResponseMessageType is:]<xs:complexType name=""ExportItemsResponseMessageType""><xs:complexContent><xs:extension
                    base=""m:ResponseMessageType""><xs:sequence><xs:element name=""ItemId""type=""t:ItemIdType"" minOccurs=""0"" maxOccurs=""1""/>
                    <xs:element name=""Data"" type=""xs:base64Binary"" minOccurs=""0"" maxOccurs=""1""/></xs:sequence> </xs:extension></xs:complexContent></xs:complexType>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R171");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R171
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                171,
                @"[In m:ExportItemsResponseMessageType Complex Type]The type of Data is xs:base64Binary ([XMLSCHEMA2]).");
        }

        /// <summary>
        /// This method is used to verify the base type ResponseMessageType related requirements.
        /// </summary>
        /// <param name="resMessage">Specified ResponseMessageType instance</param>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyResponseMessageType(ResponseMessageType resMessage, bool isSchemaValidated)
        {
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1434");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1434
            Site.CaptureRequirementIfIsTrue(
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

            if (resMessage.ResponseCodeSpecified == true)
            {
                // Verify the ResponseCodeType related requirements.
                this.VerifyResponseCodeType(isSchemaValidated);
            }

            // Verify the ResponseClassType related requirements.
            this.VerifyResponseClassType(isSchemaValidated);
        }

        /// <summary>
        /// This method is used to verify the ServerVersion related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyServerVersion(bool isSchemaValidated)
        {
            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1339");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1339
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    1339,
                    @"[In t:ServerVersionInfo Element] 
                <xs:element name=""t:ServerVersionInfo"">
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
        }

        /// <summary>
        /// This method is used to verify the ResponseCodeType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyResponseCodeType(bool isSchemaValidated)
        {
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R197");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R197
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                197,
                @"[In m:ResponseCodeType Simple Type] The type [ResponseCodeType] is defined as follow:
                <xs:simpleType name=""ResponseCodeType"">
                    <xs:restriction base=""xs:string"">
                        <xs:enumeration value=""NoError""/>
                        <xs:enumeration value=""ErrorAccessDenied""/>
                        <xs:enumeration value=""ErrorAccessModeSpecified""/>
                        <xs:enumeration value=""ErrorAccountDisabled""/>
                        <xs:enumeration value=""ErrorAddDelegatesFailed""/>
                        <xs:enumeration value=""ErrorAddressSpaceNotFound""/>
                        <xs:enumeration value=""ErrorADOperation""/>
                        <xs:enumeration value=""ErrorADSessionFilter""/>
                        <xs:enumeration value=""ErrorADUnavailable""/>
                        <xs:enumeration value=""ErrorAffectedTaskOccurrencesRequired""/>
                        <xs:enumeration value=""ErrorArchiveFolderPathCreation""/>
                        <xs:enumeration value=""ErrorArchiveMailboxNotEnabled""/>
                        <xs:enumeration value=""ErrorArchiveMailboxServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorAvailabilityConfigNotFound""/>
                        <xs:enumeration value=""ErrorBatchProcessingStopped""/>
                        <xs:enumeration value=""ErrorCalendarCannotMoveOrCopyOccurrence""/>
                        <xs:enumeration value=""ErrorCalendarCannotUpdateDeletedItem""/>
                        <xs:enumeration value=""ErrorCalendarCannotUseIdForOccurrenceId""/>
                        <xs:enumeration value=""ErrorCalendarCannotUseIdForRecurringMasterId""/>
                        <xs:enumeration value=""ErrorCalendarDurationIsTooLong""/>
                        <xs:enumeration value=""ErrorCalendarEndDateIsEarlierThanStartDate""/>
                        <xs:enumeration value=""ErrorCalendarFolderIsInvalidForCalendarView""/>
                        <xs:enumeration value=""ErrorCalendarInvalidAttributeValue""/>
                        <xs:enumeration value=""ErrorCalendarInvalidDayForTimeChangePattern""/>
                        <xs:enumeration value=""ErrorCalendarInvalidDayForWeeklyRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarInvalidPropertyState""/>
                        <xs:enumeration value=""ErrorCalendarInvalidPropertyValue""/>
                        <xs:enumeration value=""ErrorCalendarInvalidRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarInvalidTimeZone""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForTentative""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForTentative""/>
                        <xs:enumeration value=""ErrorCalendarIsNotOrganizer""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForTentative""/>
                        <xs:enumeration
                             value=""ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange""/>
                        <xs:enumeration value=""ErrorCalendarOccurrenceIsDeletedFromRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarOutOfRange""/>
                        <xs:enumeration value=""ErrorCalendarMeetingRequestIsOutOfDate""/>
                        <xs:enumeration value=""ErrorCalendarViewRangeTooBig""/>
                        <xs:enumeration value=""ErrorCallerIsInvalidADAccount""/>
                        <xs:enumeration value=""ErrorCannotArchiveCalendarContactTaskFolderException""/>
                        <xs:enumeration value=""ErrorCannotArchiveItemsInPublicFolders""/>
                        <xs:enumeration value=""ErrorCannotArchiveItemsInArchiveMailbo""/>
                        <xs:enumeration value=""ErrorCannotCreateCalendarItemInNonCalendarFolder""/>
                        <xs:enumeration value=""ErrorCannotCreateContactInNonContactFolder""/>
                        <xs:enumeration value=""ErrorCannotCreatePostItemInNonMailFolder""/>
                        <xs:enumeration value=""ErrorCannotCreateTaskInNonTaskFolder""/>
                        <xs:enumeration value=""ErrorCannotDeleteObject""/>
                        <xs:enumeration value=""ErrorCannotDisableMandatoryExtension""/>
                        <xs:enumeration value=""ErrorCannotGetSourceFolderPath""/>
                        <xs:enumeration value=""ErrorCannotGetExternalEcpUrl""/>
                        <xs:enumeration value=""ErrorCannotOpenFileAttachment""/>
                        <xs:enumeration value=""ErrorCannotDeleteTaskOccurrence""/>
                        <xs:enumeration value=""ErrorCannotEmptyFolder""/>
                        <xs:enumeration 
                            value=""ErrorCannotSetCalendarPermissionOnNonCalendarFolder""/>
                        <xs:enumeration 
                            value=""ErrorCannotSetNonCalendarPermissionOnCalendarFolder""/>
                        <xs:enumeration value=""ErrorCannotSetPermissionUnknownEntries""/>
                        <xs:enumeration value=""ErrorCannotSpecifySearchFolderAsSourceFolder""/>
                        <xs:enumeration value=""ErrorCannotUseFolderIdForItemId""/>
                        <xs:enumeration value=""ErrorCannotUseItemIdForFolderId""/>
                        <xs:enumeration value=""ErrorChangeKeyRequired""/>
                        <xs:enumeration value=""ErrorChangeKeyRequiredForWriteOperations""/>
                        <xs:enumeration value=""ErrorClientDisconnected""/>
                        <xs:enumeration value=""ErrorClientIntentInvalidStateDefinition""/>
                        <xs:enumeration value=""ErrorClientIntentNotFound""/>
                        <xs:enumeration value=""ErrorConnectionFailed""/>
                        <xs:enumeration value=""ErrorContainsFilterWrongType""/>
                        <xs:enumeration value=""ErrorContentConversionFailed""/>
                        <xs:enumeration value=""ErrorContentIndexingNotEnabled""/>
                        <xs:enumeration value=""ErrorCorruptData""/>
                        <xs:enumeration value=""ErrorCreateItemAccessDenied""/>
                        <xs:enumeration value=""ErrorCreateManagedFolderPartialCompletion""/>
                        <xs:enumeration value=""ErrorCreateSubfolderAccessDenied""/>
                        <xs:enumeration value=""ErrorCrossMailboxMoveCopy""/>
                        <xs:enumeration value=""ErrorCrossSiteRequest""/>
                        <xs:enumeration value=""ErrorDataSizeLimitExceeded""/>
                        <xs:enumeration value=""ErrorDataSourceOperation""/>
                        <xs:enumeration value=""ErrorDelegateAlreadyExists""/>
                        <xs:enumeration value=""ErrorDelegateCannotAddOwner""/>
                        <xs:enumeration value=""ErrorDelegateMissingConfiguration""/>
                        <xs:enumeration value=""ErrorDelegateNoUser""/>
                        <xs:enumeration value=""ErrorDelegateValidationFailed""/>
                        <xs:enumeration value=""ErrorDeleteDistinguishedFolder""/>
                        <xs:enumeration value=""ErrorDeleteItemsFailed""/>
                        <xs:enumeration value=""ErrorDeleteUnifiedMessagingPromptFailed""/>
                        <xs:enumeration value=""ErrorDistinguishedUserNotSupported""/>
                        <xs:enumeration value=""ErrorDistributionListMemberNotExist""/>
                        <xs:enumeration value=""ErrorDuplicateInputFolderNames""/>
                        <xs:enumeration value=""ErrorDuplicateUserIdsSpecified""/>
                        <xs:enumeration value=""ErrorEmailAddressMismatch""/>
                        <xs:enumeration value=""ErrorEventNotFound""/>
                        <xs:enumeration value=""ErrorExceededConnectionCount""/>
                        <xs:enumeration value=""ErrorExceededSubscriptionCount""/>
                        <xs:enumeration value=""ErrorExceededFindCountLimit""/>
                        <xs:enumeration value=""ErrorExpiredSubscription""/>
                        <xs:enumeration value=""ErrorExtensionNotFound""/>
                        <xs:enumeration value=""ErrorFolderCorrupt""/>
                        <xs:enumeration value=""ErrorFolderNotFound""/>
                        <xs:enumeration value=""ErrorFolderPropertRequestFailed""/>
                        <xs:enumeration value=""ErrorFolderSave""/>
                        <xs:enumeration value=""ErrorFolderSaveFailed""/>
                        <xs:enumeration value=""ErrorFolderSavePropertyError""/>
                        <xs:enumeration value=""ErrorFolderExists""/>
                        <xs:enumeration value=""ErrorFreeBusyGenerationFailed""/>
                        <xs:enumeration value=""ErrorGetServerSecurityDescriptorFailed""/>
                        <xs:enumeration value=""ErrorImContactLimitReached""/>
                        <xs:enumeration value=""ErrorImGroupDisplayNameAlreadyExists""/>
                        <xs:enumeration value=""ErrorImGroupLimitReached""/>
                        <xs:enumeration value=""ErrorImpersonateUserDenied""/>
                        <xs:enumeration value=""ErrorImpersonationDenied""/>
                        <xs:enumeration value=""ErrorImpersonationFailed""/>
                        <xs:enumeration value=""ErrorIncorrectSchemaVersion""/>
                        <xs:enumeration value=""ErrorIncorrectUpdatePropertyCount""/>
                        <xs:enumeration value=""ErrorIndividualMailboxLimitReached""/>
                        <xs:enumeration value=""ErrorInsufficientResources""/>
                        <xs:enumeration value=""ErrorInternalServerError""/>
                        <xs:enumeration value=""ErrorInternalServerTransientError""/>
                        <xs:enumeration value=""ErrorInvalidAccessLevel""/>
                        <xs:enumeration value=""ErrorInvalidArgument""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentId""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentSubfilter""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentSubfilterTextFilter""/>
                        <xs:enumeration value=""ErrorInvalidAuthorizationContext""/>
                        <xs:enumeration value=""ErrorInvalidChangeKey""/>
                        <xs:enumeration value=""ErrorInvalidClientSecurityContext""/>
                        <xs:enumeration value=""ErrorInvalidCompleteDate""/>
                        <xs:enumeration value=""ErrorInvalidContactEmailAddress""/>
                        <xs:enumeration value=""ErrorInvalidContactEmailIndex""/>
                        <xs:enumeration value=""ErrorInvalidCrossForestCredentials""/>
                        <xs:enumeration value=""ErrorInvalidDelegatePermission""/>
                        <xs:enumeration value=""ErrorInvalidDelegateUserId""/>
                        <xs:enumeration value=""ErrorInvalidExcludesRestriction""/>
                        <xs:enumeration value=""ErrorInvalidExpressionTypeForSubFilter""/>
                        <xs:enumeration value=""ErrorInvalidExtendedProperty""/>
                        <xs:enumeration value=""ErrorInvalidExtendedPropertyValue""/>
                        <xs:enumeration value=""ErrorInvalidFolderId""/>
                        <xs:enumeration value=""ErrorInvalidFolderTypeForOperation""/>
                        <xs:enumeration value=""ErrorInvalidFractionalPagingParameters""/>
                        <xs:enumeration value=""ErrorInvalidFreeBusyViewType""/>
                        <xs:enumeration value=""ErrorInvalidId""/>
                        <xs:enumeration value=""ErrorInvalidIdEmpty""/>
                        <xs:enumeration value=""ErrorInvalidIdMalformed""/>
                        <xs:enumeration value=""ErrorInvalidIdMalformedEwsLegacyIdFormat""/>
                        <xs:enumeration value=""ErrorInvalidIdMonikerTooLong""/>
                        <xs:enumeration value=""ErrorInvalidIdNotAnItemAttachmentId""/>
                        <xs:enumeration value=""ErrorInvalidIdReturnedByResolveNames""/>
                        <xs:enumeration value=""ErrorInvalidIdStoreObjectIdTooLong""/>
                        <xs:enumeration value=""ErrorInvalidIdTooManyAttachmentLevels""/>
                        <xs:enumeration value=""ErrorInvalidIdXml""/>
                        <xs:enumeration value=""ErrorInvalidImContactId""/>
                        <xs:enumeration value=""ErrorInvalidImDistributionGroupSmtpAddress""/>
                        <xs:enumeration value=""ErrorInvalidImGroupId""/>
                        <xs:enumeration value=""ErrorInvalidIndexedPagingParameters""/>
                        <xs:enumeration value=""ErrorInvalidInternetHeaderChildNodes""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationArchiveItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCreateItemAttachment""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCreateItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationAcceptItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationDeclineItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCancelItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationExpandDL""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationRemoveItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationSendItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationTentative""/>
                        <xs:enumeration value=""ErrorInvalidLogonType""/>
                        <xs:enumeration value=""ErrorInvalidMailbox""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderProperty""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderQuota""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderSize""/>
                        <xs:enumeration value=""ErrorInvalidMergedFreeBusyInterval""/>
                        <xs:enumeration value=""ErrorInvalidNameForNameResolution""/>
                        <xs:enumeration value=""ErrorInvalidOperation""/>
                        <xs:enumeration value=""ErrorInvalidNetworkServiceContext""/>
                        <xs:enumeration value=""ErrorInvalidOofParameter""/>
                        <xs:enumeration value=""ErrorInvalidPagingMaxRows""/>
                        <xs:enumeration value=""ErrorInvalidParentFolder""/>
                        <xs:enumeration value=""ErrorInvalidPercentCompleteValue""/>
                        <xs:enumeration value=""ErrorInvalidPermissionSettings""/>
                        <xs:enumeration value=""ErrorInvalidPhoneCallId""/>
                        <xs:enumeration value=""ErrorInvalidPhoneNumber""/>
                        <xs:enumeration value=""ErrorInvalidUserInfo""/>
                        <xs:enumeration value=""ErrorInvalidPropertyAppend""/>
                        <xs:enumeration value=""ErrorInvalidPropertyDelete""/>
                        <xs:enumeration value=""ErrorInvalidPropertyForExists""/>
                        <xs:enumeration value=""ErrorInvalidPropertyForOperation""/>
                        <xs:enumeration value=""ErrorInvalidPropertyRequest""/>
                        <xs:enumeration value=""ErrorInvalidPropertySet""/>
                        <xs:enumeration value=""ErrorInvalidPropertyUpdateSentMessage""/>
                        <xs:enumeration value=""ErrorInvalidProxySecurityContext""/>
                        <xs:enumeration value=""ErrorInvalidPullSubscriptionId""/>
                        <xs:enumeration value=""ErrorInvalidPushSubscriptionUrl""/>
                        <xs:enumeration value=""ErrorInvalidRecipients""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilter""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterComparison""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterOrder""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterTextFilter""/>
                        <xs:enumeration value=""ErrorInvalidReferenceItem""/>
                        <xs:enumeration value=""ErrorInvalidRequest""/>
                        <xs:enumeration value=""ErrorInvalidRestriction""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagTypeMismatch""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagInvisible""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagIdGuid""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagInheritance""/>
                        <xs:enumeration value=""ErrorInvalidRoutingType""/>
                        <xs:enumeration value=""ErrorInvalidScheduledOofDuration""/>
                        <xs:enumeration value=""ErrorInvalidSchemaVersionForMailboxVersion""/>
                        <xs:enumeration value=""ErrorInvalidSecurityDescriptor""/>
                        <xs:enumeration value=""ErrorInvalidSendItemSaveSettings""/>
                        <xs:enumeration value=""ErrorInvalidSerializedAccessToken""/>
                        <xs:enumeration value=""ErrorInvalidServerVersion""/>
                        <xs:enumeration value=""ErrorInvalidSid""/>
                        <xs:enumeration value=""ErrorInvalidSIPUri""/>
                        <xs:enumeration value=""ErrorInvalidSmtpAddress""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterType""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterTypeNotAttendeeType""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterTypeNotRecipientType""/>
                        <xs:enumeration value=""ErrorInvalidSubscription""/>
                        <xs:enumeration value=""ErrorInvalidSubscriptionRequest""/>
                        <xs:enumeration value=""ErrorInvalidSyncStateData""/>
                        <xs:enumeration value=""ErrorInvalidTimeInterval""/>
                        <xs:enumeration value=""ErrorInvalidUserOofSettings""/>
                        <xs:enumeration value=""ErrorInvalidUserPrincipalName""/>
                        <xs:enumeration value=""ErrorInvalidUserSid""/>
                        <xs:enumeration value=""ErrorInvalidUserSidMissingUPN""/>
                        <xs:enumeration value=""ErrorInvalidValueForProperty""/>
                        <xs:enumeration value=""ErrorInvalidWatermark""/>
                        <xs:enumeration value=""ErrorIPGatewayNotFound""/>
                        <xs:enumeration value=""ErrorIrresolvableConflict""/>
                        <xs:enumeration value=""ErrorItemCorrupt""/>
                        <xs:enumeration value=""ErrorItemNotFound""/>
                        <xs:enumeration value=""ErrorItemPropertyRequestFailed""/>
                        <xs:enumeration value=""ErrorItemSave""/>
                        <xs:enumeration value=""ErrorItemSavePropertyError""/>
                        <xs:enumeration value=""ErrorLegacyMailboxFreeBusyViewTypeNotMerged""/>
                        <xs:enumeration value=""ErrorLocalServerObjectNotFound""/>
                        <xs:enumeration value=""ErrorLogonAsNetworkServiceFailed""/>
                        <xs:enumeration value=""ErrorMailboxConfiguration""/>
                        <xs:enumeration value=""ErrorMailboxDataArrayEmpty""/>
                        <xs:enumeration value=""ErrorMailboxDataArrayTooBig""/>
                        <xs:enumeration value=""ErrorMailboxHoldNotFound""/>
                        <xs:enumeration value=""ErrorMailboxLogonFailed""/>
                        <xs:enumeration value=""ErrorMailboxMoveInProgress""/>
                        <xs:enumeration value=""ErrorMailboxStoreUnavailable""/>
                        <xs:enumeration value=""ErrorMailRecipientNotFound""/>
                        <xs:enumeration value=""ErrorMailTipsDisabled""/>
                        <xs:enumeration value=""ErrorManagedFolderAlreadyExists""/>
                        <xs:enumeration value=""ErrorManagedFolderNotFound""/>
                        <xs:enumeration value=""ErrorManagedFoldersRootFailure""/>
                        <xs:enumeration value=""ErrorMeetingSuggestionGenerationFailed""/>
                        <xs:enumeration value=""ErrorMessageDispositionRequired""/>
                        <xs:enumeration value=""ErrorMessageSizeExceeded""/>
                        <xs:enumeration value=""ErrorMimeContentConversionFailed""/>
                        <xs:enumeration value=""ErrorMimeContentInvalid""/>
                        <xs:enumeration value=""ErrorMimeContentInvalidBase64String""/>
                        <xs:enumeration value=""ErrorMissingArgument""/>
                        <xs:enumeration value=""ErrorMissingEmailAddress""/>
                        <xs:enumeration value=""ErrorMissingEmailAddressForManagedFolder""/>
                        <xs:enumeration value=""ErrorMissingInformationEmailAddress""/>
                        <xs:enumeration value=""ErrorMissingInformationReferenceItemId""/>
                        <xs:enumeration value=""ErrorMissingItemForCreateItemAttachment""/>
                        <xs:enumeration value=""ErrorMissingManagedFolderId""/>
                        <xs:enumeration value=""ErrorMissingRecipients""/>
                        <xs:enumeration value=""ErrorMissingUserIdInformation""/>
                        <xs:enumeration value=""ErrorMoreThanOneAccessModeSpecified""/>
                        <xs:enumeration value=""ErrorMoveCopyFailed""/>
                        <xs:enumeration value=""ErrorMoveDistinguishedFolder""/>
                        <xs:enumeration value=""ErrorMultiLegacyMailboxAccess""/>
                        <xs:enumeration value=""ErrorNameResolutionMultipleResults""/>
                        <xs:enumeration value=""ErrorNameResolutionNoMailbox""/>
                        <xs:enumeration value=""ErrorNameResolutionNoResults""/>
                        <xs:enumeration value=""ErrorNoApplicableProxyCASServersAvailable""/>
                        <xs:enumeration value=""ErrorNoCalendar""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToKerberosRequirements""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToSSLRequirements""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToVersionMismatch""/>
                        <xs:enumeration value=""ErrorNoFolderClassOverride""/>
                        <xs:enumeration value=""ErrorNoFreeBusyAccess""/>
                        <xs:enumeration value=""ErrorNonExistentMailbox""/>
                        <xs:enumeration value=""ErrorNonPrimarySmtpAddress""/>
                        <xs:enumeration value=""ErrorNoPropertyTagForCustomProperties""/>
                        <xs:enumeration value=""ErrorNoPublicFolderReplicaAvailable""/>
                        <xs:enumeration value=""ErrorNoPublicFolderServerAvailable""/>
                        <xs:enumeration value=""ErrorNoRespondingCASInDestinationSite""/>
                        <xs:enumeration value=""ErrorNotDelegate""/>
                        <xs:enumeration value=""ErrorNotEnoughMemory""/>
                        <xs:enumeration value=""ErrorObjectTypeChanged""/>
                        <xs:enumeration value=""ErrorOccurrenceCrossingBoundary""/>
                        <xs:enumeration value=""ErrorOccurrenceTimeSpanTooBig""/>
                        <xs:enumeration value=""ErrorOperationNotAllowedWithPublicFolderRoot""/>
                        <xs:enumeration value=""ErrorParentFolderIdRequired""/>
                        <xs:enumeration value=""ErrorParentFolderNotFound""/>
                        <xs:enumeration value=""ErrorPasswordChangeRequired""/>
                        <xs:enumeration value=""ErrorPasswordExpired""/>
                        <xs:enumeration value=""ErrorPhoneNumberNotDialable""/>
                        <xs:enumeration value=""ErrorPropertyUpdate""/>
                        <xs:enumeration value=""ErrorPromptPublishingOperationFailed""/>
                        <xs:enumeration value=""ErrorPropertyValidationFailure""/>
                        <xs:enumeration value=""ErrorProxiedSubscriptionCallFailure""/>
                        <xs:enumeration value=""ErrorProxyCallFailed""/>
                        <xs:enumeration value=""ErrorProxyGroupSidLimitExceeded""/>
                        <xs:enumeration value=""ErrorProxyRequestNotAllowed""/>
                        <xs:enumeration value=""ErrorProxyRequestProcessingFailed""/>
                        <xs:enumeration value=""ErrorProxyServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorProxyTokenExpired""/>
                        <xs:enumeration value=""ErrorPublicFolderMailboxDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderOperationFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderRequestProcessingFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderServerNotFound""/>
                        <xs:enumeration value=""ErrorPublicFolderSyncException""/>
                        <xs:enumeration value=""ErrorQueryFilterTooLong""/>
                        <xs:enumeration value=""ErrorQuotaExceeded""/>
                        <xs:enumeration value=""ErrorReadEventsFailed""/>
                        <xs:enumeration value=""ErrorReadReceiptNotPending""/>
                        <xs:enumeration value=""ErrorRecurrenceEndDateTooBig""/>
                        <xs:enumeration value=""ErrorRecurrenceHasNoOccurrence""/>
                        <xs:enumeration value=""ErrorRemoveDelegatesFailed""/>
                        <xs:enumeration value=""ErrorRequestAborted""/>
                        <xs:enumeration value=""ErrorRequestStreamTooBig""/>
                        <xs:enumeration value=""ErrorRequiredPropertyMissing""/>
                        <xs:enumeration value=""ErrorResolveNamesInvalidFolderType""/>
                        <xs:enumeration value=""ErrorResolveNamesOnlyOneContactsFolderAllowed""/>
                        <xs:enumeration value=""ErrorResponseSchemaValidation""/>
                        <xs:enumeration value=""ErrorRestrictionTooLong""/>
                        <xs:enumeration value=""ErrorRestrictionTooComplex""/>
                        <xs:enumeration value=""ErrorResultSetTooBig""/>
                        <xs:enumeration value=""ErrorInvalidExchangeImpersonationHeaderData""/>
                        <xs:enumeration value=""ErrorSavedItemFolderNotFound""/>
                        <xs:enumeration value=""ErrorSchemaValidation""/>
                        <xs:enumeration value=""ErrorSearchFolderNotInitialized""/>
                        <xs:enumeration value=""ErrorSendAsDenied""/>
                        <xs:enumeration value=""ErrorSendMeetingCancellationsRequired""/>
                        <xs:enumeration 
                            value=""ErrorSendMeetingInvitationsOrCancellationsRequired""/>
                        <xs:enumeration value=""ErrorSendMeetingInvitationsRequired""/>
                        <xs:enumeration value=""ErrorSentMeetingRequestUpdate""/>
                        <xs:enumeration value=""ErrorSentTaskRequestUpdate""/>
                        <xs:enumeration value=""ErrorServerBusy""/>
                        <xs:enumeration value=""ErrorServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorStaleObject""/>
                        <xs:enumeration value=""ErrorSubmissionQuotaExceeded""/>
                        <xs:enumeration value=""ErrorSubscriptionAccessDenied""/>
                        <xs:enumeration value=""ErrorSubscriptionDelegateAccessNotSupported""/>
                        <xs:enumeration value=""ErrorSubscriptionNotFound""/>
                        <xs:enumeration value=""ErrorSubscriptionUnsubscribed""/>
                        <xs:enumeration value=""ErrorSyncFolderNotFound""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotFound""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotLinkedToSharePoint""/>
                        <xs:enumeration value=""ErrorTeamMailboxUrlValidationFailed""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotAuthorizedOwner""/>
                        <xs:enumeration value=""ErrorTeamMailboxActiveToPendingDelete""/>
                        <xs:enumeration value=""ErrorTeamMailboxFailedSendingNotifications""/>
                        <xs:enumeration value=""ErrorTeamMailboxErrorUnknown""/>
                        <xs:enumeration value=""ErrorTimeIntervalTooBig""/>
                        <xs:enumeration value=""ErrorTimeoutExpired""/>
                        <xs:enumeration value=""ErrorTimeZone""/>
                        <xs:enumeration value=""ErrorToFolderNotFound""/>
                        <xs:enumeration value=""ErrorTokenSerializationDenied""/>
                        <xs:enumeration value=""ErrorUpdatePropertyMismatch""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingDialPlanNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingReportDataNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingPromptNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingRequestFailed""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingServerNotFound""/>
                        <xs:enumeration value=""ErrorUnableToGetUserOofSettings""/>
                        <xs:enumeration value=""ErrorUnableToRemoveImContactFromGroup""/>
                        <xs:enumeration value=""ErrorUnsupportedSubFilter""/>
                        <xs:enumeration value=""ErrorUnsupportedCulture""/>
                        <xs:enumeration value=""ErrorUnsupportedMapiPropertyType""/>
                        <xs:enumeration value=""ErrorUnsupportedMimeConversion""/>
                        <xs:enumeration value=""ErrorUnsupportedPathForQuery""/>
                        <xs:enumeration value=""ErrorUnsupportedPathForSortGroup""/>
                        <xs:enumeration value=""ErrorUnsupportedPropertyDefinition""/>
                        <xs:enumeration value=""ErrorUnsupportedQueryFilter""/>
                        <xs:enumeration value=""ErrorUnsupportedRecurrence""/>
                        <xs:enumeration value=""ErrorUnsupportedTypeForConversion""/>
                        <xs:enumeration value=""ErrorUpdateDelegatesFailed""/>
                        <xs:enumeration value=""ErrorUserNotUnifiedMessagingEnabled""/>
                        <xs:enumeration value=""ErrorValueOutOfRange""/>
                        <xs:enumeration value=""ErrorVoiceMailNotImplemented""/>
                        <xs:enumeration value=""ErrorVirusDetected""/>
                        <xs:enumeration value=""ErrorVirusMessageDeleted""/>
                        <xs:enumeration value=""ErrorWebRequestInInvalidState""/>
                        <xs:enumeration value=""ErrorWin32InteropError""/>
                        <xs:enumeration value=""ErrorWorkingHoursSaveFailed""/>
                        <xs:enumeration value=""ErrorWorkingHoursXmlMalformed""/>
                        <xs:enumeration value=""ErrorWrongServerVersion""/>
                        <xs:enumeration value=""ErrorWrongServerVersionDelegate""/>
                        <xs:enumeration value=""ErrorMissingInformationSharingFolderId""/>
                        <xs:enumeration value=""ErrorDuplicateSOAPHeader""/>
                        <xs:enumeration value=""ErrorSharingSynchronizationFailed""/>
                        <xs:enumeration value=""ErrorSharingNoExternalEwsAvailable""/>
                        <xs:enumeration value=""ErrorFreeBusyDLLimitReached""/>
                        <xs:enumeration value=""ErrorInvalidGetSharingFolderRequest""/>
                        <xs:enumeration value=""ErrorNotAllowedExternalSharingByPolicy""/>
                        <xs:enumeration value=""ErrorUserNotAllowedByPolicy""/>
                        <xs:enumeration value=""ErrorPermissionNotAllowedByPolicy""/>
                        <xs:enumeration value=""ErrorOrganizationNotFederated""/>
                        <xs:enumeration value=""ErrorMailboxFailover""/>
                        <xs:enumeration value=""ErrorInvalidExternalSharingInitiator""/>
                        <xs:enumeration value=""ErrorMessageTrackingPermanentError""/>
                        <xs:enumeration value=""ErrorMessageTrackingTransientError""/>
                        <xs:enumeration value=""ErrorMessageTrackingNoSuchDomain""/>
                        <xs:enumeration value=""ErrorUserWithoutFederatedProxyAddress""/>
                        <xs:enumeration value=""ErrorInvalidOrganizationRelationshipForFreeBusy""/>
                        <xs:enumeration value=""ErrorInvalidFederatedOrganizationId""/>
                        <xs:enumeration value=""ErrorInvalidExternalSharingSubscriber""/>
                        <xs:enumeration value=""ErrorInvalidSharingData""/>
                        <xs:enumeration value=""ErrorInvalidSharingMessage""/>
                        <xs:enumeration value=""ErrorNotSupportedSharingMessage""/>
                        <xs:enumeration value=""ErrorApplyConversationActionFailed""/>
                        <xs:enumeration value=""ErrorInboxRulesValidationError""/>
                        <xs:enumeration value=""ErrorOutlookRuleBlobExists""/>
                        <xs:enumeration value=""ErrorRulesOverQuota""/>
                        <xs:enumeration value=""ErrorNewEventStreamConnectionOpened""/>
                        <xs:enumeration value=""ErrorMissedNotificationEvents""/>
                        <xs:enumeration value=""ErrorDuplicateLegacyDistinguishedName""/>
                        <xs:enumeration value=""ErrorInvalidClientAccessTokenRequest""/>
                        <xs:enumeration value=""ErrorNoSpeechDetected""/>
                        <xs:enumeration value=""ErrorUMServerUnavailable""/>
                        <xs:enumeration value=""ErrorRecipientNotFound""/>
                        <xs:enumeration value=""ErrorRecognizerNotInstalled""/>
                        <xs:enumeration value=""ErrorSpeechGrammarError""/>
                        <xs:enumeration value=""ErrorInvalidManagementRoleHeader""/>
                        <xs:enumeration value=""ErrorLocationServicesDisabled""/>
                        <xs:enumeration value=""ErrorLocationServicesRequestTimedOut""/>
                        <xs:enumeration value=""ErrorLocationServicesRequestFailed""/>
                        <xs:enumeration value=""ErrorLocationServicesInvalidRequest""/>
                        <xs:enumeration value=""ErrorMailboxScopeNotAllowedWithoutQueryString""/>
                        <xs:enumeration value=""ErrorArchiveMailboxSearchFailed""/>
                        <xs:enumeration value=""ErrorArchiveMailboxServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorInvalidPhotoSize""/>
                        <xs:enumeration value=""ErrorSearchQueryHasTooManyKeywords""/>
                        <xs:enumeration value=""ErrorSearchTooManyMailboxes""/>
                        <xs:enumeration value=""ErrorDiscoverySearchesDisabled""/>
                    </xs:restriction>
                </xs:simpleType>");
        }

        /// <summary>
        /// This method is used to verify the ResponseClassType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyResponseClassType(bool isSchemaValidated)
        {
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R191");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R191
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                191,
                @"[In t:ResponseClassType Simple Type] The type [ResponseClassType] is defined as follow:
                    <xs:simpleType name=""ResponseClassType"">
                      <xs:restriction
                        base=""xs:string""
                      >
                        <xs:enumeration
                          value=""Error""
                         />
                        <xs:enumeration
                          value=""Success""
                         />
                        <xs:enumeration
                          value=""Warning""
                         />
                      </xs:restriction>
                    </xs:simpleType>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1436");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1436
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1436,
                @"[In m:ResponseMessageType Complex Type] [ResponseClass:] The following values are valid for this attribute:
                Success,
                Warning,
                Error.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1284");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1284
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1284,
                @"[In m:ResponseMessageType Complex Type] This attribute [ResponseClass] MUST be present.");
        }

        /// <summary>
        /// This method is used to verify the UploadItemsResponseMessageType related requirements.
        /// </summary>
        /// <param name="uploadItemsMessage">Specified UploadItemsResponseMessageType instance</param>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyUploadItemsResponseMessageType(UploadItemsResponseMessageType uploadItemsMessage, bool isSchemaValidated)
        {
            // verify the base type ResponseMessageType related requirements.
            this.VerifyResponseMessageType(uploadItemsMessage as ResponseMessageType, isSchemaValidated);

            // If the schema validation and the above base type verification are successful, 
            // then MS-OXWSBTRF_R127, MS-OXWSBTRF_R199, MS-OXWSBTRF_R201 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R127");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R127
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                127,
                @"[In m:UploadItemsResponseMessageType Complex Type][The schema of UploadItemsResponseMessageType is:]<xs:complexType name=""UploadItemsResponseMessageType"">
                  <xs:complexContent>
                    <xs:extension base=""m:ResponseMessageType"">
                      <xs:sequence>
                        <xs:element name=""ItemId""  type=""t:ItemIdType"" minOccurs=""0""/>
                      </xs:sequence>
                    </xs:extension>
                  </xs:complexContent>
                </xs:complexType>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R199");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R199
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                199,
                @"[In m:UploadItemsResponseMessageType Complex Type]The Type of ItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSBTRF_R201");

            // Verify MS-OXWSBTRF requirement: MS-OXWSBTRF_R201
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                201,
                @"[In m:UploadItemsResponseMessageType Complex Type]Only a single instance of this element [ItemId] can be present.");
        }

        /// <summary>
        /// The method is used to verify the related base BaseResponseMessageType requirements.
        /// </summary>
        /// <param name="isSchemaValidated">Whether the schema is validated.</param>
        private void VerifyBaseResponseMessageType(bool isSchemaValidated)
        {
            // If the schema validation is successful, then MS-OXWSCDATA_R1091, MS-OXWSCDATA_R1092 and MS-OXWSCDATA_R1036 can be captured
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1091");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1091
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1091,
                @"[In m:BaseResponseMessageType Complex Type] The BaseResponseMessageType complex type MUST NOT be sent in a SOAP message because it is an abstract type.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1092");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1092
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1092,
                @"[In m:BaseResponseMessageType Complex Type]The type [BaseResponseMessageType] is defined as follow:<xs:complexType name=""BaseResponseMessageType"">
                  <xs:sequence>
                    <xs:element name=""ResponseMessages"" type=""m:ArrayOfResponseMessagesType""/>
                  </xs:sequence>
                </xs:complexType>");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1036");

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

            // The schema defines  <xs:element name="ResponseMessages" type="m:ArrayOfResponseMessagesType"/>, maxoccurs default value is 1.
            // Then if the schema validation is successful, then element ResponseMessages maxoccurs MUST be one. Therefore MS-OXWSCDATA_R1094 can be captured.
            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1094");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1094
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1094,
                @"[In m:BaseResponseMessageType Complex Type] There MUST be only one ResponseMessages element in a response.");
        }
    }
}