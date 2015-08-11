namespace Microsoft.Protocols.TestSuites.MS_ADMINS
{
    using System;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The capture requirement part of adapter's implementation
    /// </summary>
    public partial class MS_ADMINSAdapter : ManagedAdapterBase, IMS_ADMINSAdapter
    {
        /// <summary>
        /// Verify the requirements of the transport when the response is received successfully.
        /// </summary>
        private void VerifyTransportRelatedRequirements()
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

                    if (Common.IsRequirementEnabled(1002, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS transport is supported. If the HTTPS transport is not supported, the response can't be received successfully.
                        Site.CaptureRequirement(
                        1002,
                        @"[In Transport]Implementation does additionally support SOAP over HTTPS for securing communication with clients.(Windows SharePoint Services 3.0 and above products follow this behavior.)");
                    }

                    break;

                default:
                    Site.Debug.Fail("Unknown transport type " + transport);

                    break;
            }

            // Add the log information.
            Site.Log.Add(Microsoft.Protocols.TestTools.LogEntryKind.Comment, "The protocol message is formatted as version: {0}", this.adminService.SoapVersion);

            // Verify MS-ADMINS requirement: MS-ADMINS_R4.
            bool isR4Verified = this.adminService.SoapVersion == SoapProtocolVersion.Soap11 || this.adminService.SoapVersion == SoapProtocolVersion.Soap12;
            Site.CaptureRequirementIfIsTrue(
                isR4Verified,
                4,
                @" [In Transport]Protocol messages MUST be formatted as specified either in [SOAP1.1], section 4 ""SOAP Envelope"" or in [SOAP1.2/1], section 5 ""SOAP Message Construct.""");
            
            // Having received the response successfully have proved the XML Schema is used. If the structures do not use XML Schema, the response can't be received successfully.
            Site.CaptureRequirement(
                6,
                @"[In Common Message Syntax]The syntax of the definitions uses XML schema as defined in [XMLSCHEMA1] and [XMLSCHEMA2], and WSDL as defined in [WSDL].");
        }          
    
        /// <summary>
        /// Verify the requirements of the SOAP fault when the SOAP fault is received.
        /// </summary>
        /// <param name="soapExp">The returned SOAP fault</param>
        private void VerifySoapFaultRequirements(SoapException soapExp) 
        {
            // If a SOAP fault is returned and the SOAP fault is not null, which means protocol server faults are returned using SOAP faults, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                soapExp,
                3000,
                @" [In Transport]Protocol server faults MUST be returned using SOAP faults as specified either in [SOAP1.1], section 4.4 ""SOAP Fault"" or in [SOAP1.2/1], section 5.4 ""SOAP Fault.""");

            // The schemas are validated during invoking the operations and an XmlSchemaValidationException will be thrown if anyone of the schemas is incorrect, there is no such exception thrown, so the following requirement can be captured.
            Site.CaptureRequirement(
                7,
                @"[In SOAPFaultDetails]The schema of SOAPFaultDetails is defined as: <s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap"">
                   <s:complexType name=""SOAPFaultDetails"">
                      <s:sequence>
                         <s:element name=""errorstring"" type=""s:string""/>
                         <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
                      </s:sequence>
                   </s:complexType>
                </s:schema>");
           
            // Extract the errorcode from SOAP fault.
            string strErrorCode = Common.ExtractErrorCodeFromSoapFault(soapExp);
            if (strErrorCode != null)
            {
                Site.Assert.IsTrue(strErrorCode.StartsWith("0x", StringComparison.CurrentCultureIgnoreCase), "The error code value should start with '0x'.");
                Site.Assert.IsTrue(strErrorCode.Length == 10, "The error code value's length should be 10.");
    
                // If the value of the errorcode starts with "0x" and the length is 10, which means the format of the value is 0xAAAAAAAA, then the following requirement can be captured.
                Site.CaptureRequirement(
                    2015,
                    @"[In SOAPFaultDetails]The format inside the string [errorcode] MUST be 0xAAAAAAAA.");
            }

            // If a SOAP fault is returned and the SOAP fault is not null, which means protocol server faults are returned using SOAP faults, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                soapExp,
                2019,
                @"[In Protocol Details]This protocol [MS-ADMINS] allows protocol servers to notify protocol clients of application-level faults using SOAP faults.");

            // The schemas are validated during invoking the operations and an XmlSchemaValidationException will be thrown if anyone of the schemas is incorrect, there is no such exception thrown and only SOAPFaultDetails complex type is included in the schemas, which means the detail element in the SOAP faults is conforms to the SOAPFaultDetails complex type, so the following requirement can be captured.
            Site.CaptureRequirement(
                2020,
                @"[In Protocol Details]This protocol [MS-ADMINS] allows protocol servers to provide additional details for SOAP faults by including a detail element as specified either in [SOAP1.1], section 4.4 ""SOAP Fault"" or [SOAP1.2/1], section 5.4 ""SOAP Fault"" that conforms to the XML schema of the SOAPFaultDetails complex type specified in section 2.2.4.1.");
        }

        /// <summary>
        /// Validate CreateSite's response data CreateSiteResult when the response is received successfully.
        /// </summary>
        /// <param name="createSiteResult">CreateSite's response data.</param>
        private void ValidateCreateSiteResponseData(string createSiteResult)
        {
            // If the response data is not null, which means the response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createSiteResult, 
                2033,
                @"[In CreateSite][The schema of CreateSite operation is defined as:] <wsdl:operation name=""CreateSite"">
                    <wsdl:input message=""tns:CreateSiteSoapIn"" />
                    <wsdl:output message=""tns:CreateSiteSoapOut"" />
                </wsdl:operation>");

            // If the response data is not null, which means the response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createSiteResult,
                12,
                @"[In CreateSite]The client sends a CreateSiteSoapIn request message, the server responds with a CreateSiteSoapOut response message.");

            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                81,
                @"[In CreateSiteSoapOut]The SOAP body contains a CreateSiteResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2063,
                @"[In CreateSiteResponse][The schema of CreateSiteResponse element is defined as:] <s:element name=""CreateSiteResponse"">
                    <s:complexType>
                    <s:sequence>
                        <s:element minOccurs=""0"" maxOccurs=""1"" name=""CreateSiteResult"" type=""s:string""/>
                    </s:sequence>
                    </s:complexType>
                </s:element>");

            // If the response have been received successfully and the response is a valid URL, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                Uri.IsWellFormedUriString(createSiteResult, UriKind.Absolute),
                1077,
                @"[In CreateSiteResponse]CreateSiteResult: Specifies the URL of the new site collection.");
        }

        /// <summary>
        /// Validate DeleteSite's response when the response is received successfully.
        /// </summary>
        private void ValidateDeleteSiteResponse()
        {
            // The response have been received successfully, then the following requirement can be captured. 
            // If the following requirement is fail, the response can't be received successfully.
            Site.CaptureRequirement(
                2065,
                @"[In DeleteSite][The schema of DeleteSite operation is defined as:] <wsdl:operation name=""DeleteSite"">
                        <wsdl:input message=""tns:DeleteSiteSoapIn"" />
                        <wsdl:output message=""tns:DeleteSiteSoapOut"" />
                    </wsdl:operation>");

            // The response have been received successfully, then the following requirement can be captured. 
            // If the following requirement is fail, the response can't be received successfully.
            Site.CaptureRequirement(
                82,
                @"[In DeleteSite]The client sends a DeleteSiteSoapIn request message and the server responds with a DeleteSiteSoapOut response message.");

            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                92,
                @"[In DeleteSiteSoapOut]The SOAP body contains a DeleteSiteResponse element.");

            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2073,
                @"[In DeleteSiteResponse][The schema of DeleteSiteResponse element is defined as:] 
                    <s:element name=""DeleteSiteResponse"">  <s:complexType/></s:element>");
        }

        /// <summary>
        /// Validate GetLanguages response data getLanguagesResult when the response is received successfully.
        /// </summary>
        /// <param name="getLanguagesResult">The return value of GetLanguages operation.</param>
        private void ValidateGetLanguagesResponseData(GetLanguagesResponseGetLanguagesResult getLanguagesResult)
        {
            // If the response data is not null, which means the response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                getLanguagesResult,
                2075,
                @"[In GetLanguages][The schema of GetLanguages operation is defined as:] <wsdl:operation name=""GetLanguages"">
                        <wsdl:input message=""tns:GetLanguagesSoapIn"" />
                        <wsdl:output message=""tns:GetLanguagesSoapOut"" />
                    </wsdl:operation>");

            // If the response data is not null, which means the response have been received successfully, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                getLanguagesResult,
                2082,
                @"[In GetLanguagesResponse][The schema of DeleteSiteResponse element is defined as:] <s:element name=""GetLanguagesResponse"">
                    <s:complexType>
                    <s:sequence>
                        <s:element minOccurs=""1"" maxOccurs=""1"" name=""GetLanguagesResult"">
                            <s:complexType>
                            <s:sequence>
                            <s:element name=""Languages"">
                                <s:complexType>
                                <s:sequence>
                                    <s:element maxOccurs=""unbounded"" name=""LCID"" type=""s:int"" />
                                </s:sequence>
                                </s:complexType>
                            </s:element>
                            </s:sequence>
                        </s:complexType>
                        </s:element>
                    </s:sequence>
                    </s:complexType>
                </s:element>");

            // The response have been received successfully, then the following requirement can be captured. 
            // If the following requirement is failed, the response can't be received successfully.
            Site.CaptureRequirement(
                93,
                @"[In GetLanguages]The client sends a GetLanguagesSoapIn request message and the server responds with a GetLanguagesSoapOut response message.");

            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                98,
                @"[In GetLanguagesSoapOut]The SOAP body contains a GetLanguagesResponse element.");

            // If the lcid list in the response is not empty, then the following requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                getLanguagesResult.Languages.Length > 0,
                94,
                @"[In GetLanguagesResponse]GetLanguagesResult: Provides the locale identifiers (LCIDs) of languages used in the deployment.");
        }
    }
}