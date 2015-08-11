namespace Microsoft.Protocols.TestSuites.MS_AUTHWS
{
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Adapter requirements capture code for MS_AUTHWSAdapter.
    /// </summary>
    public partial class MS_AUTHWSAdapter
    {
        /// <summary>
        /// This method is used to verify the requirements about the transport.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:

                    // As response successfully returned, the transport related requirements can be directly captured.
                    Site.CaptureRequirement(
                        1,
                        @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
                    break;

                case TransportProtocol.HTTPS:
                    if (Common.IsRequirementEnabled(121, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                            121,
                            @"[In Appendix B: Product Behavior] Implementation does additionally support SOAP over HTTPS to help secure communication with protocol clients.(The Windows SharePoint Services 3.0 and above products follow this behavior.)");
                    }

                    break;

                default:

                    Site.Debug.Fail("Unknown transport type " + transport);
                    break;
            }

            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);
            switch (soapVersion)
            {
                case SoapVersion.SOAP11:

                    // As response successfully returned, the SOAP1.1 related requirements can be directly captured.
                    Site.CaptureRequirement(
                        122,
                        @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] section 4 [or [SOAP1.2/1] section 5].");

                    break;

                case SoapVersion.SOAP12:

                    // As response successfully returned, the SOAP1.2 related requirements can be directly captured.
                    Site.CaptureRequirement(
                        123,
                        @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.2/1] section 5.");

                    break;

                default:
                    Site.Debug.Fail("Unknown soap version" + soapVersion);
                    break;
            }
        }

        /// <summary>
        /// Validate common message syntax to schema and capture related requirements.
        /// </summary>
        private void ValidateAndCaptureCommonMessageSyntax()
        {
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // If the server response is validated successfully, we can make sure that the server responds with an correct message syntax, MS-AUTHWS_R7, MS-AUTHWS_R8 and MS-AUTHWS_R9 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                7,
                @"[In Common Message Syntax] The syntax of the definitions uses the XML Schema, as specified in [XMLSCHEMA1] and [XMLSCHEMA2].");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                8,
                @"[In Common Message Syntax] The syntax of the definitions uses Web Services Description Language, as specified in [WSDL].");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                9,
                @"[In Namespaces] This specification[MS-AUTHWS] defines and references various XML namespaces using the mechanisms specified in [XMLNS].");
        }

        /// <summary>
        /// Validate the Login Response.
        /// </summary>
        /// <param name="result">The response result of Login.</param>
        private void ValidateLoginResponse(LoginResult result)
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // If the server response is validated successfully, we can make sure that the server responds with an LoginSoapOut response message, MS-AUTHWS_49, MS-AUTHWS_51, MS-AUTHWS_57 and MS-AUTHWS_40 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                49,
                @"[In Login] [The Login operation is defined as follow:]
<wsdl:operation name=""Login"">
    <wsdl:input message=""tns:LoginSoapIn"" />
    <wsdl:output message=""tns:LoginSoapOut"" />
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                51,
                @"[In Login] [If the protocol client sends a LoginSoapIn request WSDL message] and the protocol server responds with a LoginSoapOut response WSDL message, as specified in section 3.1.4.1.1.2.");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                57,
                @"[In LoginSoapOut] The LoginSoapOut message is the response WSDL message that is used by a protocol server when logging on a user in response to a LoginSoapIn request message.");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                40,
                @"[In Message Processing Events and Sequencing Rules] The following table summarizes the list of WSDL operations[Login, Mode] that are defined by this protocol.");

            // If the server response is validated successfully, and the LoginResponse has returned, MS-AUTHWS_59 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "LoginResponse"),
                59,
                @"[In LoginSoapOut] The SOAP body contains a LoginResponse element, as specified in section 3.1.4.1.2.2.");

            // If the server response is validated successfully, we can make sure that the server responds with LoginResponse, MS-AUTHWS_66 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                66,
                @"[In LoginResponse] [The LoginResponse element is defined as follows:]
<s:element name=""LoginResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""LoginResult"" type=""tns:LoginResult""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            if (result != null)
            {
                // If LoginResult element exist, and the server response pass the validation successfully, we can make sure LoginResult is defined according to the schema, MS-AUTHWS_71,  MS-AUTHWS_67, MS-AUTHWS_69 can be verified. 
                Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                71,
                @"[In LoginResult] [The LoginResult complex type is defined as follows:]
<s:complexType name=""LoginResult"">
  <s:sequence>
    <s:element name=""CookieName"" type=""s:string"" minOccurs=""0""/>
    <s:element name=""ErrorCode"" type=""tns:LoginErrorCode""/>
    <s:element name=""TimeoutSeconds"" type=""s:int"" minOccurs=""0"" maxOccurs=""1""/>
  </s:sequence>
</s:complexType>");

                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    67,
                    @"[In LoginResponse] LoginResult: A LoginResult complex type, as specified in section 3.1.4.1.3.1.");

                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    69,
                    @"[In LoginResult] The LoginResult complex type contains an error code.");

                // If LoginResult element exist, and the server response pass the validation successfully, we can make sure LoginErrorCode is defined according to the schema, MS-AUTHWS_79, MS-AUTHWS_75 and MS-AUTHWS_80 can be verified. 
                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    79,
                    @"[In LoginErrorCode] [The LoginErrorCode simple type is defined as follows:]
<s:simpleType name=""LoginErrorCode"">
  <s:restriction base=""s:string"">
    <s:enumeration value=""NoError""/>
    <s:enumeration value=""NotInFormsAuthenticationMode""/>
    <s:enumeration value=""PasswordNotMatch""/>
  </s:restriction>
</s:simpleType>");

                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    75,
                    @"[In LoginResult] ErrorCode: An error code, as specified in section 3.1.4.1.4.1.");

                Site.CaptureRequirementIfIsTrue(
                    isResponseValid,
                    80,
                    @"[In LoginErrorCode] The LoginErrorCode field has three allowable values:  [NoError, NotInFormsAuthenticationMode, PasswordNotMatch].");
            }

            this.ValidateAndCaptureCommonMessageSyntax();
        }

        /// <summary>
        /// Validate the Mode Response.
        /// </summary>
        private void ValidateModeResponse()
        {
            XmlElement xmlResponse = SchemaValidation.LastRawResponseXml;
            bool isResponseValid = SchemaValidation.ValidationResult == ValidationResult.Success;

            // If the server response is validated successfully, we can make sure that the server responds with an ModeSoapOut response message, MS-AUTHWS_87, MS-AUTHWS_89, MS-AUTHWS_95 and MS-AUTHWS_40 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                87,
                @"[In Mode] [The Mode operation is defined as follows:]
<wsdl:operation name=""Mode"">
    <wsdl:input message=""tns:ModeSoapIn"" />
    <wsdl:output message=""tns:ModeSoapOut"" />
</wsdl:operation>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                89,
                @"[In Mode] [If the protocol client sends a ModeSoapIn request WSDL message] and the protocol server responds with a ModeSoapOut response WSDL message.");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                95,
                @"[In ModeSoapOut] The ModeSoapOut message is the response WSDL message that a protocol server sends after retrieving the authentication mode.");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                40,
                @"[In Message Processing Events and Sequencing Rules] The following table summarizes the list of WSDL operations[Login, Mode] that are defined by this protocol.");

            // If the server response is validated successfully, and the ModeResponse has returned, MS-AUTHWS_97 can be verified.
            Site.CaptureRequirementIfIsTrue(
                this.ResponseExists(xmlResponse, "ModeResponse"),
                97,
                @"[In ModeSoapOut] The SOAP body contains a ModeResponse element, as specified in section 3.1.4.2.2.2.");

            // If the server response is validated successfully, we can make sure that the server responds with ModeResponse, MS-AUTHWS_102 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                102,
                @"[In ModeResponse] [The ModeResponse element is defined as follows:]
<s:element name=""ModeResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""ModeResult"" type=""tns:AuthenticationMode""/>
    </s:sequence>
  </s:complexType>
</s:element>");

            // If ModeResult element exist, and the server response pass the validation successfully, we can make sure LoginResult is defined according to the schema, MS-AUTHWS_106, MS-AUTHWS_103 and MS-AUTHWS_107 can be verified. 
            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                106,
                @"[In AuthenticationMode] [The AuthenticationMode simple type is defined as follows:]
<s:simpleType name=""AuthenticationMode"">
  <s:restriction base=""s:string"">
    <s:enumeration value=""None""/>
    <s:enumeration value=""Windows""/>
    <s:enumeration value=""Passport""/>
    <s:enumeration value=""Forms""/>
  </s:restriction>
</s:simpleType>");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                103,
                @"[In ModeResponse] ModeResult: An AuthenticationMode simple type, as specified in section 3.1.4.2.4.1.");

            Site.CaptureRequirementIfIsTrue(
                isResponseValid,
                107,
                @"[In AuthenticationMode] The AuthenticationMode field has four allowable values:  [None, Windows, Passport, Forms].");

            this.ValidateAndCaptureCommonMessageSyntax();
        }

        /// <summary>
        /// This method is used to validate whether there is an element named as responseName in xmlElement.
        /// </summary>
        /// <param name="xmlElement">A XmlElement object which is the raw XML response from server.</param>
        /// <param name="responseName">The name of the response element.</param>
        /// <returns>A Boolean value indicates whether there is an element named as responseName.</returns>
        private bool ResponseExists(XmlElement xmlElement, string responseName)
        {
            // The first child is the response element.
            XmlNode firstChildNode = xmlElement.ChildNodes[0].FirstChild;
            return firstChildNode.Name == responseName;
        }
    }
}