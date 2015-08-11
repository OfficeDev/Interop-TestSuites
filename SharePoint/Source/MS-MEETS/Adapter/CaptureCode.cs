namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Adapter requirements capture code for MS-MEETSAdapter server role.
    /// </summary>
    public partial class MS_MEETSAdapter
    {
        /// <summary>
        /// Capture underlying transport protocol related requirements.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:

                    // As response successfully returned, the transport related requirements can be captured.
                    Site.CaptureRequirement(
                        1,
                        @"[In Transport]Protocol servers MUST support SOAP over HTTP.");
                    break;

                case TransportProtocol.HTTPS:

                    if (Common.IsRequirementEnabled(3020, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                        3020,
                        @"[In Appendix B: Product Behavior]Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients.(The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                    }

                    break;

                default:
                    Site.Debug.Fail("Unknown transport type " + transport);
                    break;
            }

            // Add the log information.
            Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "The protocol message is formatted as version: {0}", this.service.SoapVersion);

            // Verifies MS-MEETS requirement: MS-MEETS_R3.
            bool isR3Verified = this.service.SoapVersion == SoapProtocolVersion.Soap11 || this.service.SoapVersion == SoapProtocolVersion.Soap12;
            Site.CaptureRequirementIfIsTrue(
                isR3Verified,
                3,
                @"[In Transport]Protocol messages MUST be formatted as specified in [SOAP1.1]section 4, SOAP Envelope, or in [SOAP1.2/1]section 5, SOAP Message Construct.");
        }

        /// <summary>
        /// Validate SOAP Fault message according to schema and capture related requirements.
        /// </summary>
        /// <param name="exception">The SoapException thrown.</param>
        private void ValidateAndCaptureSoapFaultRequirements(SoapException exception)
        {
            // Since there is an SoapException returned, MS-MEETS_R4101 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                exception,
                4101,
                @"[In Transport][Protocol server faults MUST be returned by]using SOAP faults as specified in [SOAP1.1]section 4.4, SOAP Fault or [SOAP1.2/1]section 5.4, SOAP Fault.");

            bool isResponseValid = SchemaValidation.ValidateXml(this.Site, SchemaValidation.GetSoapFaultDetailBody(exception.Detail.OuterXml)) == ValidationResult.Success;

            // If the schema validation for SOAP fault is success, MS-MEETS_R3021 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                3021,
                @"[In Detail]Whenever an operation in this protocol fails, it [server] returns a SOAP fault in this format [as in Detail schema].");

            // If the schema validation for SOAP fault is success, MS-MEETS_R3022 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                3022,
                @"[In Detail]This element [Detail]is defined as follows.
                <s:element name=""detail"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""errorstring"" type=""s:string"" minOccurs=""1"" maxOccurs=""1""/>
                      <s:element name=""errorcode"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""/>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // After validating SOAP fault message xml, we can make sure that the response soap fault body contains them.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                10,
                @"[In Detail]It [detail]consists of a SOAP fault code combined with SOAP fault detail text that describes the error.");
        }
        
        /// <summary>
        ///  Validate common message syntax to schema and capture related requirements.
        /// </summary>
        private void ValidateAndCaptureCommonMessageSyntax()
        {
            // If the server response pass the validation successfully, we can make sure that the server responds with an correct message syntax, MS-MEETS_R8 and MS-MEETS_R2942 can be verified.
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;
        
            // Verify MS-MEETS requirement: MS-MEETS_R8.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                8,
                @"[In Common Message Syntax]The syntax of the definitions uses XML schema as defined in [XMLSCHEMA1]and [XMLSCHEMA2], and Web Services Description Language (WSDL) as specified in [WSDL].");
        
            // Verify MS-MEETS requirement: MS-MEETS_R2942.  
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                2942,
                @"[In Namespaces]This specification defines and references various XML namespaces by using the mechanisms specified in [XMLNS].");
        }

        /// <summary>
        /// Verifies the AddMeeting Response.
        /// </summary>
        /// <param name="result">The response information of AddMeeting</param>
        private void VerifyAddMeetingResponse(AddMeetingResponseAddMeetingResult result)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with an AddMeetingSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                36,
                @"[In AddMeeting]This [AddMeeting]operation is defined as follows. 
                <wsdl:operation name=""AddMeeting"">
                    <wsdl:input message=""AddMeetingSoapIn"" />
                    <wsdl:output message=""AddMeetingSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with an AddMeetingSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R38.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                38,
                @"[In AddMeeting][if the client sends an AddMeetingSoapIn request message]and the protocol server responds with an AddMeetingSoapOut response message(section 3.1.4.1.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R47.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "AddMeetingResponse"),
                47,
                @"[In AddMeetingSoapOut]The SOAP body contains an AddMeetingResponse element (section 3.1.4.1.2.2).");

            // If the server response pass the validation successfully, we can make sure that the AddMeetingResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R65.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                65,
                @"[In AddMeetingResponse]This element[AddMeetingResponse]is defined as follows. 
             <s:element name=""AddMeetingResponse"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""AddMeetingResult"" minOccurs=""0"">
                    <s:complexType mixed=""true"">
                      <s:sequence>
                        <s:element name=""AddMeeting"" type=""tns:AddMeeting""/>
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                </s:sequence>
              </s:complexType>
            </s:element>");

            // If the server response pass the validation successfully, we can make sure that AddMeeting operation used when adding meeting not based on the Gregorian calendar.
            // Verify MS-MEETS requirement: MS-MEETS_R39
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                39,
                @"[In AddMeeting]This operation [AddMeeting]MUST be used when adding meetings not based on the Gregorian calendar, because the iCalendar format as used in the AddMeetingFromICal operation (section 3.1.4.2) does not support non-Gregorian recurring meetings.");

            // If the server response pass the validation successfully, we can make sure that AddMeeting operation was not used to add a recurring meeting workspace.
            // Verify MS-MEETS requirement: MS-MEETS_R41
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                41,
                @"[In AddMeeting]This operation [AddMeeting]MUST NOT be used to add meetings to a recurring meeting workspace.");

            // If AddMeetingResult element exist, and the server response pass the validation successfully, 
            // we can make sure the AddMeeting is defined according to the schema.
            if (result != null)
            {
                this.VerifyAddMeetingComplexType(isResponseValid);
            }
        }

        /// <summary>
        /// Verifies the response of AddMeetingFromICal.
        /// </summary>
        /// <param name="result">The response information of AddMeetingFromICal</param>
        private void VerifyAddMeetingFromICalResponse(AddMeetingFromICalResponseAddMeetingFromICalResult result)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with an AddMeetingFromICalSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                68,
                @"[In AddMeetingFromICal]This operation [AddMeetingFromICal method]is defined as follows. 
                <wsdl:operation name=""AddMeetingFromICal"">
                   <wsdl:input message=""AddMeetingFromICalSoapIn"" />
                   <wsdl:output message=""AddMeetingFromICalSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with an AddMeetingFromICalSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                70,
                @"[In AddMeetingFromICal][if the client sends an AddMeetingFromICalSoapIn request message]The protocol server responds with an AddMeetingFromICalSoapOut response message (section 3.1.4.2.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R77.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "AddMeetingFromICalResponse"),
                77,
                @"[In AddMeetingFromICalSoapOut]The SOAP body contains an AddMeetingFromICalResponse element (section 3.1.4.2.2.2).");

            // If the server response pass the validation successfully, we can make sure that the AddMeetingFromICal operation does not support non-Gregorian recurring meetings.
            // Verify MS-MEETS requirement: MS-MEETS_R71.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                71,
                @"[In AddMeetingFromICal]The AddMeeting operation (section 3.1.4.1) MUST be used when adding meetings not based on the Gregorian calendar because this operation [AddMeetingFromICal]does not support non-Gregorian recurring meetings.");

            // If the server response pass the validation successfully, we can make sure that the AddMeetingFromICalResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R86.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                86,
                @"[In AddMeetingFromICalResponse]This element [AddMeetingFromICalResponse]is defined as follows. 
                    <s:element name=""AddMeetingFromICalResponse"">
                      <s:complexType>
                        <s:sequence>
                          <s:element name=""AddMeetingFromICalResult"" minOccurs=""0"">
                            <s:complexType mixed=""true"">
                              <s:sequence>
                                <s:element name=""AddMeetingFromICal"" type=""tns:AddMeetingFromICal"" />
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                        </s:sequence>
                      </s:complexType>
                    </s:element>");

            if (result != null)
            {
                // If AddMeetingFromICalResult element exist, and the server response pass the validation successfully, 
                // we can make sure AddMeetingFromICal is defined according to the schema.
                // Verifies MS-MEETS requirement: MS-MEETS_R8601.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    8601,
                    @"[In AddMeetingFromICalResponse] [AddMeetingFromICal complex type is defined as follows: ]
                    <s:complexType name=""AddMeetingFromICal"">
                      <s:complexContent>
                        <s:extension base=""tns:AddMeeting"">
                          <s:sequence>
                            <s:element name=""AttendeeUpdateStatus"" type=""tns:AttendeeUpdateStatus"" />
                          </s:sequence>
                        </s:extension>
                      </s:complexContent>
                    </s:complexType>");

                // If AddMeetingFromICalResult element exist, and the server response pass the validation successfully, 
                // we can make sure AddMeetingFromICal is defined according to the schema.
                this.VerifyAddMeetingComplexType(isResponseValid);

                // If AddMeetingFromICalResult element exist, and the server response pass the validation successfully, 
                // we can make sure AttendeeUpdateStatus is defined according to the schema.
                this.VerifyAttendeeUpdateStatusComplexType(isResponseValid);

                // If the server response pass the validation successfully, MS-MEETS_R87 can be verified.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    87,
                    @"[In AddMeetingFromICalResponse]AddMeetingFromICalResult: The response XML consists of two elements containing information about the meeting instance newly-added to the Meeting Workspace.");
            }
        }

        /// <summary>
        /// Verifies the response of CreateWorkspace.
        /// </summary>
        /// <param name="result">The response information of CreateWorkspace</param>
        private void VerifyCreateWorkspaceResponse(CreateWorkspaceResponseCreateWorkspaceResult result)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a CreateWorkspaceSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                3001,
                @"[In CreateWorkspace]This operation [CreateWorkspace]is defined as follows.
                <wsdl:operation name=""CreateWorkspace"">
                    <wsdl:input message=""CreateWorkspaceSoapIn"" />
                    <wsdl:output message=""CreateWorkspaceSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a CreateWorkspaceSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R92.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                92,
                @"[In CreateWorkspace][if client sends a CreateWorkspaceSoapIn request message]The protocol server responds with a CreateWorkspaceSoapOut response message (section 3.1.4.3.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R99.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "CreateWorkspaceResponse"),
                99,
                @"[In CreateWorkspaceSoapOut]The SOAP body contains a CreateWorkspaceResponse element (section 3.1.4.3.2.2).");

            // If the server response pass the validation successfully, we can make sure that the CreateWorkspaceResponse is defined according to the schema.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                110,
                 @"[In CreateWorkspaceResponse]This element [CreateWorkspaceResponse]is defined as follows. 
                 <s:element name=""CreateWorkspaceResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""CreateWorkspaceResult"" minOccurs=""0"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""CreateWorkspace"" 
                                       type=""tns:CreateWorkspace""/>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            if (result != null)
            {
                // If CreateWorkspaceResult element exist, and the server response pass the validation successfully, 
                // we can make sure CreateWorkspace is defined according to the schema. MS-MEETS_R11001 and MS-MEETS_R111 can be verified.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    11001,
                    @"[In CreateWorkspaceResponse] [CreateWorkspace complex type is defined as follows: ]
                    <s:complexType name=""CreateWorkspace"">
                      <s:attribute name=""Url"" type=""s:string""/>
                    </s:complexType>");

                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    111,
                    @"[In CreateWorkspaceResponse]CreateWorkspaceResult: The response XML consists of one element containing information about the newly created meeting workspace.");
            }
        }

        /// <summary>
        /// The method is used to verify the response of DeleteWorkspace.
        /// </summary>       
        private void VerifyDeleteWorkspaceResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a DeleteWorkspaceSoapOut response message. MS-MEETS_R151, MS-MEETS_R153, MS-MEETS_R161 and MS-MEETS_R159 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                151,
                @"[In DeleteWorkspace]This operation[DeleteWorkspace]is defined as follows. 
                <wsdl:operation name=""DeleteWorkspace"">
                    <wsdl:input message=""DeleteWorkspaceSoapIn"" />
                    <wsdl:output message=""DeleteWorkspaceSoapOut"" />
                </wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                153,
                @"[In DeleteWorkspace][if the client sends a DeleteWorkspaceSoapIn request message]and the protocol server responds with a DeleteWorkspaceSoapOut response message (section 3.1.4.4.1.2).");

            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "DeleteWorkspaceResponse"),
                161,
                @"[In DeleteWorkspaceSoapOut]The SOAP body contains a DeleteWorkspaceResponse element (section 3.1.4.4.2.2).");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                159,
                @"[In DeleteWorkspaceSoapOut]The DeleteWorkspaceSoapOut message contains the response for the DeleteWorkspace operation (section 3.1.4.4).");

            // If the server response pass the validation successfully, we can make sure that the DeleteWorkspaceResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R164.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                164,
                @"[In DeleteWorkspaceResponse]This element[DeleteWorkspaceResponse]is defined as follows. 
                <s:element name=""DeleteWorkspaceResponse"">
                 <s:complexType/>
                </s:element>");
        }

        /// <summary>
        /// Verifies the response of GetMeetingsInformation.
        /// </summary>
        /// <param name="result">The response information of GetMeetingsInformation</param>
        /// <param name="requestFlags">The flag of the request</param>
        private void VerifyGetMeetingsInformationResponse(GetMeetingsInformationResponseGetMeetingsInformationResult result, MeetingInfoTypes? requestFlags)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a GetMeetingsInformationSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                168,
                @"[In GetMeetingsInformation]This operation[GetMeetingsInformation] is defined as follows. 
               <wsdl:operation name=""GetMeetingsInformation"">
                <wsdl:input message=""GetMeetingsInformationSoapIn"" />
                <wsdl:output message=""GetMeetingsInformationSoapOut"" />
               </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a GetMeetingsInformationSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R170.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                170,
                @"[In GetMeetingsInformation][if the client sends a GetMeetingsInformationSoapIn request message]and the protocol server responds with a GetMeetingsInformationSoapOut response message (section 3.1.4.5.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R176.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "GetMeetingsInformationResponse"),
                176,
                @"[In GetMeetingsInformationSoapOut]The SOAP body contains a GetMeetingsInformationResponse element  (section 3.1.4.5.2.2).");

            // If the server response pass the validation successfully, we can make sure that the GetMeetingsInformationResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R188.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                188,
                 @"[In GetMeetingsInformationResponse]This element [GetMeetingsInformationResponse]is defined as follows.
                    <s:element name=""GetMeetingsInformationResponse"">
                      <s:complexType>
                        <s:sequence>
                          <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetMeetingsInformationResult"">
                            <s:complexType mixed=""true"">
                              <s:sequence>
                                <s:element name=""MeetingsInformation"">
                                  <s:complexType mixed=""true"">
                                    <s:element minOccurs=""0"" maxOccurs=""1"" name=""AllowCreate"" type=""tns:CaseInsensitiveTrueFalse"" />
                                    <s:element name=""ListTemplateLanguages"" minOccurs=""0"">
                                      <s:complexType>
                                        <s:sequence>
                                         <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""LCID"" type=""s:string"" />
                                         </s:sequence>
                                       </s:complexType>
                                     </s:element>
                                     <s:element name=""ListTemplates"" minOccurs=""0"">
                                       <s:complexType>
                                        <s:sequence>
                                          <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""Template"" type=""tns:Template"" />
                                        </s:sequence>
                                      </s:complexType>
                                    </s:element>
                                    <s:element minOccurs=""0"" maxOccurs=""1"" name=""WorkspaceStatus"" type=""tns:WorkspaceStatus"" />                
                                  </s:complexType>
                                </s:element>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                        </s:sequence>
                      </s:complexType>
                    </s:element>");

            // If the GetMeetingsInformationResult element exist, check the following requirements.
            if (result != null)
            {
                // Check response xml content and requestFlags parameter relationship and capture related requirements
                if (requestFlags.HasValue && (requestFlags & MeetingInfoTypes.AllowCreate) != 0)
                {
                    // If AllowCreate is exist and server response pass the validation successfully, MS-MEETS_R192 can be verified.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        192,
                         @"[In GetMeetingsInformationResponse][AllowCreate is defined as follows:]
                    <s:element name=""AllowCreate"" type=""CaseInsensitiveTrueFalse"" minOccurs=""0""/>");

                    // If server response pass the validation successfully, MS-MEETS_R2931 can be verified.                    
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        2931,
                        @"[In CaseInsensitiveTrueFalse]The definition of the CaseInsensitiveTrueFalse simple type is as follows:
                    <s:simpleType name=""CaseInsensitiveTrueFalse"">
                      <s:restriction base=""s:string"">
                        <s:pattern value=""[Tt][Rr][Uu][Ee]|[Ff][Aa][Ll][Ss][Ee]""/>
                      </s:restriction>
                    </s:simpleType>");
                }

                if (requestFlags.HasValue && (requestFlags & MeetingInfoTypes.QueryLanguages) != 0)
                {
                    // Verifies MS-MEETS requirement: MS-MEETS_R3041
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        3041,
                       @"[In GetMeetingsInformationResponse][ListTemplateLanguages is defined as follows:]
                        <s:element name=""ListTemplateLanguages"" minOccurs=""0"">
                          <s:complexType>
                            <s:sequence>
                              <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""LCID"" type=""s:string""/>
                            </s:sequence>
                          </s:complexType>
                        </s:element>");
                }

                if (requestFlags.HasValue && (requestFlags & MeetingInfoTypes.QueryTemplates) != 0)
                {
                    // Verifies MS-MEETS requirement: MS-MEETS_R201
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        201,
                         @"[In GetMeetingsInformationResponse][ListTemplates is defined as:]
                    <s:element name=""ListTemplates"" minOccurs=""0"">
                      <s:complexType>
                        <s:sequence>
                          <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""Template"" type=""tns:Template""/>
                        </s:sequence>
                      </s:complexType>
                    </s:element>");

                    if (this.ElementExists(SchemaValidation.LastRawResponseXml, "Template"))
                    {
                        // If Template element exist, and the server response pass the validation successfully, 
                        // we can make sure Template is defined according to the schema.
                        Site.CaptureRequirementIfIsTrue(
                            isResponseValid,
                            20101,
                            @"[In GetMeetingsInformationResponse][Template complex type is defined as follows:]
                                <s:complexType name=""Template"">
                                  <s:attribute name=""Name"" type=""s:string""/>
                                  <s:attribute name=""Title"" type=""s:string""/>
                                  <s:attribute name=""Id"" type=""s:int""/>
                                  <s:attribute name=""Description"" type=""s:string""/>
                                  <s:attribute name=""ImageUrl"" type=""s:string""/>
                                </s:complexType>");
                    }
                }

                if (requestFlags.HasValue && (requestFlags & MeetingInfoTypes.QueryOthers) != 0)
                {
                    // Verifies MS-MEETS requirement: MS-MEETS_R212.
                    Site.CaptureRequirementIfIsTrue(
                        this.ElementExists(SchemaValidation.LastRawResponseXml, "WorkspaceStatus"),
                        212,
                        @"[In GetMeetingsInformationResponse]This element [WorkspaceStatus]is present in the response when bit flag 0x8 is specified inrequestFlags.");

                   // Verifies MS-MEETS requirement: MS-MEETS_R21201.
                  Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    21201,
                    @"[In GetMeetingsInformationResponse][WorkspaceStatus element is defined as follows:]
                      <s:element name=""WorkspaceStatus"" type=""tns:WorkspaceStatus"" minOccurs=""0""/>
                        <s:complexType name=""WorkspaceStatus"">
                           <s:attribute name=""UniquePermissions"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                           <s:attribute name=""MeetingCount"" type=""tns:UnsignedIntOrEmpty""/>
                           <s:attribute name=""AnonymousAccess"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                           <s:attribute name=""AllowAuthenticatedUsers"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                        </s:complexType>
                       </s:element>");

                    // Verifies MS-MEETS requirement: MS-MEETS_R18801.          
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        18801,
                         @"[In GetMeetingsInformationResponse][WorkspaceStatus complex type is defined as follows:]
                        <s:complexType name=""WorkspaceStatus"">
                          <s:attribute name=""UniquePermissions"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                          <s:attribute name=""MeetingCount"" type=""tns:UnsignedIntOrEmpty""/>
                          <s:attribute name=""AnonymousAccess"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                          <s:attribute name=""AllowAuthenticatedUsers"" type=""tns:CaseInsensitiveTrueFalseOrEmpty""/>
                        </s:complexType>");

                    // Verifies MS-MEETS requirement: MS-MEETS_R2932.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        2932,
                        @"[In CaseInsensitiveTrueFalseOrEmpty]The definition of the CaseInsensitiveTrueFalseOrEmpty simple type is as follows:
                    <s:simpleType name=""CaseInsensitiveTrueFalseOrEmpty"">
                      <s:restriction base=""s:string"">
                        <s:pattern value=""([Tt][Rr][Uu][Ee]|[Ff][Aa][Ll][Ss][Ee])?""/>
                      </s:restriction>
                    </s:simpleType>");

                    // Verifies MS-MEETS requirement: MS-MEETS_R2933.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        2933,
                        @"[In UnsignedIntOrEmpty]The definition of the UnsignedIntOrEmpty simple type is as follows:
                    <s:simpleType name=""UnsignedIntOrEmpty"">
                      <s:union memberTypes=""s:unsignedInt tns:Empty""/>
                    </s:simpleType>");

                    // Verifies MS-MEETS requirement: MS-MEETS_R2934.                       
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        2934,
                        @"[In Empty]The definition of the Empty simple type is as follows.
                        <s:simpleType name=""Empty"">
                           <s:restriction base=""s:string"">
                             <s:maxLength value=""0""/>
                           </s:restriction>
                         </s:simpleType>");
                }

                // Since the situation of requestFlag set to 0x1(AllowCreate),0x2(QueryLanguages),0x4(QueryTemplates),0x8(QueryOthers) have verified in above capture code, so MS-MEETS_R191 can be captured directly. 
            Site.CaptureRequirement(
                191,
                @"[In GetMeetingsInformationResponse]An element [AllowCreate or ListTemplateLanguages or ListTemplates or WorkspaceStatus]is present, if its corresponding requestFlags was set as follows.");
            }
        }

        /// <summary>
        /// Verifies the response of GetMeetingWorkspaces.
        /// </summary>
        /// <param name="result">The response information of GetMeetingWorkspaces</param>
        private void VerifyGetMeetingWorkspacesResponse(GetMeetingWorkspacesResponseGetMeetingWorkspacesResult result)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a GetMeetingWorkspacesSoapOut response message, MS-MEETS_R218 and MS-MEETS_R4008 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                218,
                @"[In GetMeetingWorkspaces]This operation [GetMeetingWorkspaces]is defined as follows. 
                <wsdl:operation name=""GetMeetingWorkspaces"">
                <wsdl:input message=""GetMeetingWorkspacesSoapIn"" />
                <wsdl:output message=""GetMeetingWorkspacesSoapOut"" />
                </wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                4008,
                @"[In GetMeetingWorkspaces][If the protocol client sends a GetMeetingWorkspacesSoapIn request message], and the protocol server responds with a GetMeetingWorkspacesSoapOut response message section (3.1.4.6.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R226.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "GetMeetingWorkspacesResponse"),
                226,
                @"[In GetMeetingWorkspacesSoapOut]The SOAP body contains a GetMeetingWorkspacesResponse element (section 3.1.4.6.2.2).");

            // If the server response pass the validation successfully, we can make sure that the GetMeetingWorkspacesResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R232.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                232,
                  @"[In GetMeetingWorkspacesResponse]This element[GetMeetingWorkspacesResponse]is defined as follows.
                 <s:element name=""GetMeetingWorkspacesResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""GetMeetingWorkspacesResult"" minOccurs=""0"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""MeetingWorkspaces"" minOccurs=""0"">
                              <s:complexType>
                                <s:sequence>
                                  <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""Workspace"" type=""tns:Workspace""/>
                                </s:sequence>
                              </s:complexType>
                            </s:element>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            if (result != null)
            {
                // If the GetMeetingWorkspacesResult element exist, and the server response pass the validation successfully, 
                // we can make sure GetMeetingWorkspacesResult is defined according to the schema.
                // Verifies MS-MEETS requirement: MS-MEETS_R233.
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    233,
                    @"[In GetMeetingWorkspacesResponse]GetMeetingWorkspacesResult: The response XML consists of one element containing a list of meeting workspaces.");

                if (this.ElementExists(SchemaValidation.LastRawResponseXml, "Workspace"))
                {
                    // If the Workspace element exist, and the server response pass the validation successfully, 
                    // we can make sure Workspace is defined according to the schema.
                    // Verifies MS-MEETS requirement: MS-MEETS_R23201.
                    Site.CaptureRequirementIfIsTrue(
                        isResponseValid,
                        23201,
                        @"[In GetMeetingWorkspacesResponse][Workspace complex type is defined as follows:]
                    <s:complexType name=""Workspace"">
                      <s:attribute name=""Url"" type=""s:string""/>
                      <s:attribute name=""Title"" type=""s:string""/>
                    </s:complexType>");
                }
            }
        }

        /// <summary>
        /// Verifies the response of RemoveMeeting.
        /// </summary>
        private void VerifyRemoveMeetingResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a RemoveMeetingSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                239,
                @"[In RemoveMeeting]This operation [RemoveMeeting]is defined as follows.
                <wsdl:operation name=""RemoveMeeting"">
                    <wsdl:input message=""RemoveMeetingSoapIn"" />
                    <wsdl:output message=""RemoveMeetingSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a RemoveMeetingSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R241.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                241,
                @"[In RemoveMeeting][if the client sends a RemoveMeetingSoapIn request message]and the protocol server responds with a RemoveMeetingSoapOut response message (section 3.1.4.7.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R248.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "RemoveMeetingResponse"),
                248,
                @"[In RemoveMeetingSoapOut]The SOAP body contains a RemoveMeetingResponse element (section 3.1.4.7.2.2).");

            // If the server response pass the validation successfully, we can make sure that the RemoveMeetingResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R258.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                258,
                 @"[In RemoveMeetingResponse]This element[RemoveMeetingResponse element]is defined as follows. 
                <s:element name=""RemoveMeetingResponse"">
                 <s:complexType/>
                </s:element>");
        }

        /// <summary>
        /// Verifies the response of RestoreMeeting.
        /// </summary>
        private void VerifyRestoreMeetingResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a RestoreMeetingSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R260.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                260,
                @"[In RestoreMeeting]This operation [RestoreMeeting]is defined as follows.
                    <wsdl:operation name=""RestoreMeeting"">
                        <wsdl:input message=""RestoreMeetingSoapIn"" />
                        <wsdl:output message=""RestoreMeetingSoapOut"" />
                    </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a RestoreMeetingSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R262.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                262,
                @"[In RestoreMeeting][if the client sends a RestoreMeetingSoapIn request message]and the protocol server responds with a RestoreMeetingSoapOut response message (section 3.1.4.8.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R270.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "RestoreMeetingResponse"),
                270,
                @"[In RestoreMeetingSoapOut]The SOAP body contains a RestoreMeetingResponse element (section 3.1.4.8.2.2).");

            // If the server response pass the validation successfully, we can make sure that the RestoreMeetingResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R273.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                273,
                @"[In RestoreMeetingResponse]This element [RestoreMeetingResponse]is defined as follow. 
                  <s:element name=""RestoreMeetingResponse""> <s:complexType/> </s:element>");
        }

        /// <summary>
        /// Verifies the response of SetAttendeeResponse.
        /// </summary>
        private void VerifySetAttendeeResponseResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a SetAttendeeResponseSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                275,
                @"[In SetAttendeeResponse]This operation [SetAttendeeResponse operation]is defined as follows. 
                <wsdl:operation name=""SetAttendeeResponse""> 
                   <wsdl:input message=""SetAttendeeResponseSoapIn"" />
                    <wsdl:output message=""SetAttendeeResponseSoapOut"" /> 
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a SetAttendeeResponseSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R277.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                277,
                @"[In SetAttendeeResponse][if the client sends a SetAttendeeResponseSoapIn request message]and the protocol server responds with a SetAttendeeResponseSoapOut response message (section 3.1.4.9.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R284.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "SetAttendeeResponseResponse"),
                284,
                @"[In SetAttendeeResponseSoapOut]The SOAP body contains a SetAttendeeResponseResponse element (section 3.1.4.9.2.2).");

            // If the server response pass the validation successfully, we can make sure that the SetAttendeeResponseResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R298.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                298,
                @"[In SetAttendeeResponseResponse]This element [SetAttendeeResponseResponse]is defined as follows.
                <s:element name=""SetAttendeeResponseResponse"">
                 <s:complexType/>
                </s:element>");
        }

        /// <summary>
        /// Verifies the response of SetWorkSpaceTitle.
        /// </summary>
        private void VerifySetWorkspaceTitleResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with a SetWorkspaceTitleSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                301,
                @"[In SetWorkspaceTitle]This[SetWorkspacetitle]operation is defined as follows. 
                <wsdl:operation name=""SetWorkspaceTitle"">
                 <wsdl:input message=""SetWorkspaceTitleSoapIn"" />
                 <wsdl:output message=""SetWorkspaceTitleSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with a SetWorkspaceTitleSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R304.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                304,
                @"[In SetWorkspaceTitle][if client sends a SetWorkspaceTitleSoapIn request message]and the protocol server responds with a SetWorkspaceTitleSoapOut response message (section 3.1.4.10.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R312.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "SetWorkspaceTitleResponse"),
                312,
                @"[In SetWorkspaceTitleSoapOut]The SOAP body contains a SetWorkspaceTitleResponse element (section 3.1.4.10.2.2).");

            // If the server response pass the validation successfully, we can make sure that the SetWorkspaceTitleResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R319.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                319,
               @"[In SetWorkspaceTitleResponse]This[SetWorkspaceTitleResponse]element is defined as follows. 
                <s:element name=""SetWorkspaceTitleResponse"">
                 <s:complexType/>
                </s:element>");
        }

        /// <summary>
        /// Verifies the response of UpdateMeeting.
        /// </summary>
        private void VerifyUpdateMeetingResponse()
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with an UpdateMeetingSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                321,
                @"[In UpdateMeeting]This[UpdateMeeting]operation is defined as follows. 
                <wsdl:operation name=""UpdateMeeting"">
                 <wsdl:input message=""UpdateMeetingSoapIn"" />
                 <wsdl:output message=""UpdateMeetingSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with an UpdateMeetingSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R323.
            Site.CaptureRequirementIfIsTrue(
               isResponseValid,
               323,
               @"[In UpdateMeeting][if client sends an UpdateMeetingSoapIn request message]The protocol server responds with an UpdateMeetingSoapOut response message (section 3.1.4.11.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R332.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "UpdateMeetingResponse"),
                332,
                @"[In UpdateMeetingSoapOut]The SOAP body contains an UpdateMeetingResponse element (section 3.1.4.11.2.2).");

            // If the server response pass the validation successfully, we can make sure that the UpdateMeetingResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R349.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                349,
                @"[In UpdateMeetingResponse]This[UpdateMeetingResponse]element is defined as follows. 
                <s:element name=""UpdateMeetingResponse"">
                 <s:complexType/>
                </s:element>");
        }

        /// <summary>
        /// Verifies the response of UpdateMeetingFromICal.
        /// </summary>
        /// <param name="result">The UpdateMeetingFromICal response information </param>
        private void VerifyUpdateMeetingFromICalResponse(UpdateMeetingFromICalResponseUpdateMeetingFromICalResult result)
        {
            bool isResponseValid = ValidationResult.Success == SchemaValidation.ValidationResult;

            // If the server response pass the validation successfully, we can make sure that the server responds with an UpdateMeetingFromICalSoapOut response message.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                352,
                @"[In UpdateMeetingFromICal]This [UpdateMeetingFromICal]operation is defined as follows. 
                <wsdl:operation name=""UpdateMeetingFromICal"">
                 <wsdl:input message=""UpdateMeetingFromICalSoapIn"" />
                 <wsdl:output message=""UpdateMeetingFromICalSoapOut"" />
                </wsdl:operation>");

            // If the server response pass the validation successfully, we can make sure that the server responds with an UpdateMeetingFromICalSoapOut response message.
            // Verifies MS-MEETS requirement: MS-MEETS_R354.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                354,
                @"[In UpdateMeetingFromICal][if client sends an UpdateMeetingFromICalSoapIn request message]the protocol server responds with an UpdateMeetingFromICalSoapOut response message (section 3.1.4.12.1.2).");

            // Verifies MS-MEETS requirement: MS-MEETS_R362.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(SchemaValidation.LastRawResponseXml, "UpdateMeetingFromICalResponse"),
                362,
                @"[In UpdateMeetingFromICalSoapOut]The SOAP body contains an UpdateMeetingFromICalResponse element (section 3.1.4.12.2.2).");

            // If the server response pass the validation successfully, we can make sure that the UpdateMeetingFromICalResponse is defined according to the schema.
            // Verifies MS-MEETS requirement: MS-MEETS_R375.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                375,
                @"[In UpdateMeetingFromICalResponse]This element[UpdateMeetingFromICalResponse]is defined as follows.
                 <s:element name=""UpdateMeetingFromICalResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""UpdateMeetingFromICalResult"" minOccurs=""0"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""UpdateMeetingFromICal"">
                              <s:complexType>
                                <s:sequence>
                                  <s:element name=""AttendeeUpdateStatus"" type=""tns:AttendeeUpdateStatus""/>
                                </s:sequence>
                              </s:complexType>
                            </s:element>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // If the UpdateMeetingFromICalResult element exist, and the server response pass the validation successfully, 
            // we can make sure AttendeeUpdateStatus is defined according to the schema.            
            if (result != null)
            {
                this.VerifyAttendeeUpdateStatusComplexType(isResponseValid);
            }
        }

        /// <summary>
        /// Verifies AddMeeting complex type.
        /// </summary>
        /// <param name="isResponseValid">Whether the server response is valid.</param>
        private void VerifyAddMeetingComplexType(bool isResponseValid)
        {
            // Verifies MS-MEETS requirement: MS-MEETS_R23
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                23,
            @"[In AddMeeting]This type [AddMeeting]is defined as follows. 
                <s:complexType name=""AddMeeting"">
                  <s:attribute name=""Url"" type=""s:string""/>
                  <s:attribute name=""HostTitle"" type=""s:string""/>
                  <s:attribute name=""UniquePermissions"" type=""s:boolean""/>
                  <s:attribute name=""MeetingCount"" type=""s:int""/>
                  <s:attribute name=""AnonymousAccess"" type=""s:boolean""/>
                  <s:attribute name=""AllowAuthenticatedUsers"" type=""s:boolean""/>
                </s:complexType>");
        }

        /// <summary>
        /// Verifies AttendeeUpdateStatus complex type.
        /// </summary>
        /// <param name="isResponseValid">Whether the server response is valid.</param>
        private void VerifyAttendeeUpdateStatusComplexType(bool isResponseValid)
        {
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                30,
                @"[In AttendeeUpdateStatus]This type [AttendeeUpdateStatus]is defined as follows. 
                 <s:complexType name=""AttendeeUpdateStatus"">
                   <s:attribute name=""Code"" type=""s:int""/>
                   <s:attribute name=""Detail"" type=""s:string""/>
                   <s:attribute name=""ManageUserPage"" type=""s:string""/>
                 </s:complexType>");
        }

        /// <summary>
        /// The method is used to validate whether there is an element Named as responseName in xmlElement.
        /// </summary>
        /// <param name="xmlElement">A XmlElement object which is the raw XML response from server.</param>
        /// <param name="responseName">The Name of the response element.</param>
        /// <returns>A Boolean value indicates whether there is an element named as responseName.</returns>
        private bool ResponseExists(XmlElement xmlElement, string responseName)
        {
            // The first child is the response element.
            XmlNode firstChildNode = xmlElement.ChildNodes[0].FirstChild;
            return firstChildNode.Name == responseName;
        }

        /// <summary>
        /// Verifies XML element exists
        /// </summary>
        /// <param name="xmlElement">A XmlElement object which is the raw XML response from server.</param>
        /// <param name="elementName">The Name of the response element</param>
        /// <returns>True if element exists, otherwise false</returns>
        private bool ElementExists(XmlElement xmlElement, string elementName)
        {
            string content = xmlElement.OuterXml;
            if (content.Contains(elementName))
            {
                return true;
            }

            return false;
        }
    }
}